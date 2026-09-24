#!/usr/bin/env python3
"""Fail-closed verifier for conditional aggregate performance receipts."""

from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any

from run_profile import BASELINE_COMMIT, CONTROL_CASES, PHASES, profile_input_snapshot, verify_candidate_freeze


HERE = Path(__file__).resolve().parent
RESULTS = HERE / "results"
BASELINE = f"baseline-{BASELINE_COMMIT}"
CANDIDATE = "candidate-final"


def expected_candidate_cases() -> tuple[str, ...]:
    # Keep this order identical to the harness. The three extra control rows
    # remain candidate-only while the nine controls below are captured from
    # both revisions.
    return (
        "scalar-control-arithmetic",
        "scalar-control-sin",
        "scalar-control-imsum",
        "database-control-dsum",
        "array-control-4x4-arithmetic",
        "array-control-4x4-sin",
        "array-control-16x16-arithmetic",
        "array-control-16x16-sin",
        "reference-array-16x4-arithmetic",
        "scalar-aggregate-sum",
        "literal-aggregate-4x4-sum",
        "reference-aggregate-64x4-sum",
        "reference-conditional-16x4-sumif-implicit",
        "reference-conditional-64x4-sumif-explicit",
        "reference-conditional-256x4-sumifs",
        "reference-conditional-1024x4-countif",
        "reference-conditional-256x4-countifs",
        "reference-conditional-64x4-averageif-implicit",
        "reference-conditional-64x4-averageif-explicit",
        "reference-conditional-1024x4-averageifs",
        "reference-conditional-16x1-countif-text",
        "reference-conditional-6x4-sumif-list",
        "reference-conditional-3d-sumif",
        "reference-conditional-2x4-sumif-anchor-clip",
        "nested-conditional-64-sumifs",
        "nested-conditional-256-sumifs",
        "nested-conditional-1024-sumifs",
        "error-conditional-averageif-empty",
        "error-conditional-averageifs-empty",
        "error-conditional-sumifs-mismatch",
        "error-conditional-countifs-mismatch",
        "error-conditional-constant-range",
        "resource-conditional-reference-cells",
    )


def load(path: Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8"))


def rows(directory: Path) -> list[dict[str, Any]]:
    path = directory / "measurements.jsonl"
    if not path.is_file():
        raise RuntimeError(f"missing measurements: {path}")
    return [json.loads(line) for line in path.read_text(encoding="utf-8").splitlines() if line.strip()]


def equal(label: str, actual: Any, expected: Any) -> None:
    if actual != expected:
        raise RuntimeError(f"{label}: expected {expected!r}, observed {actual!r}")


def conditional_read_bounds(case: str, elements: int) -> tuple[int, int] | None:
    if case == "reference-conditional-3d-sumif":
        return 5, 5
    if case == "reference-conditional-2x4-sumif-anchor-clip":
        return 10, 10
    if case.endswith("sumif-list"):
        return elements, elements
    if case == "resource-conditional-reference-cells":
        return 0, 0
    if case in {"error-conditional-averageif-empty", "error-conditional-averageifs-empty"}:
        return elements, elements
    if case.startswith("error-"):
        return None
    if case.startswith("nested-conditional-") and case.endswith("-sumifs"):
        return 11 * elements // 4, 11 * elements // 4
    if case.endswith("sumif-implicit") or case.endswith("countif") or case.endswith("countif-text"):
        return elements, elements
    if case.endswith("sumif-explicit") or case.endswith("averageif-explicit"):
        return 3 * elements // 2, 3 * elements // 2
    if case.endswith("sumifs"):
        return 7 * elements // 4, 7 * elements // 4
    if case.endswith("countifs"):
        return 3 * elements // 2, 3 * elements // 2
    if case.endswith("averageif-implicit"):
        return elements, elements
    if case.endswith("averageifs"):
        return 7 * elements // 4, 7 * elements // 4
    return None


def verify_manifest(directory: Path, profile_hashes: dict[str, str]) -> dict[str, Any]:
    manifest = load(directory / "source-manifest.json")
    before = manifest.get("before")
    after = manifest.get("after")
    if not isinstance(before, dict) or not isinstance(after, dict):
        raise RuntimeError(f"{directory}: malformed source manifest")
    for key in ("source_sha256", "workspace_source_sha256", "profile_input_sha256"):
        equal(f"{directory} {key} stable", before.get(key), after.get(key))
    equal(f"{directory} source stable", manifest.get("source_sha256_unchanged"), True)
    equal(f"{directory} workspace stable", manifest.get("workspace_source_sha256_unchanged"), True)
    equal(f"{directory} profile stable", manifest.get("profile_input_sha256_unchanged"), True)
    equal(f"{directory} profile inputs", before.get("profile_input_sha256"), profile_hashes)
    equal(f"{directory} git stable", manifest.get("git_head_unchanged"), True)
    if not before.get("harness_sha256", {}).get("Cargo.lock"):
        raise RuntimeError(f"{directory}: missing harness lock hash")
    cleanup = load(directory / "target-cleanup.json")
    equal(f"{directory} target cleanup", cleanup.get("removed"), True)
    return manifest


def verify_rows(directory: Path, records: list[dict[str, Any]], expected_cases: list[str], warmups: int, samples: int) -> None:
    expected_count = len(expected_cases) * len(PHASES) * samples
    equal(f"{directory} row count", len(records), expected_count)
    groups: dict[tuple[str, str], list[dict[str, Any]]] = {}
    for record in records:
        key = (record.get("case"), record.get("phase"))
        if key[0] not in expected_cases or key[1] not in PHASES:
            raise RuntimeError(f"{directory}: unexpected group {key}")
        groups.setdefault(key, []).append(record)
        equal(f"{directory} {key} supported", record.get("supported"), True)
        equal(f"{directory} {key} warmups", record.get("warmups"), warmups)
        if int(record.get("elapsed_ns_p50", 0)) <= 0 or int(record.get("rss_kib", 0)) <= 0:
            raise RuntimeError(f"{directory} {key}: invalid elapsed/RSS")
        if int(record.get("repeat", 0)) <= 0:
            raise RuntimeError(f"{directory} {key}: invalid repeat")
        if not record.get("validation_scope", "").startswith("one untimed direct f64 fixture oracle"):
            raise RuntimeError(f"{directory} {key}: validation scope missing")
        if not record.get("raw_stdout") or not (directory / record["raw_stdout"]).is_file():
            raise RuntimeError(f"{directory} {key}: raw stdout missing")
        if not (directory / record["raw_time"]).is_file():
            raise RuntimeError(f"{directory} {key}: raw time missing")
        raw_samples = record.get("samples")
        if not isinstance(raw_samples, list) or len(raw_samples) != 1:
            raise RuntimeError(f"{directory} {key}: child did not emit one sample")
        sample = raw_samples[0]
        for field in ("elapsed_ns", "alloc_calls", "dealloc_calls", "requested_bytes", "released_bytes", "live_before", "live_after", "peak_live_delta", "work", "memory_retained", "reference_reads"):
            if int(sample.get(field, -1)) < 0:
                raise RuntimeError(f"{directory} {key}: negative {field}")
        if sample["live_before"] != sample["live_after"]:
            raise RuntimeError(f"{directory} {key}: allocator bytes do not balance")
        if int(sample["released_bytes"]) > int(sample["requested_bytes"]):
            raise RuntimeError(f"{directory} {key}: released bytes exceed requested")
        if int(sample["checksum"]) != int(record["checksum_p50"]):
            raise RuntimeError(f"{directory} {key}: checksum receipt mismatch")
        repeat = int(record["repeat"])
        reads = int(record.get("reference_reads_p50", 0)) // repeat
        equal(f"{directory} {key} normalized reference reads", int(record.get("reference_reads_per_repeat", -1)), reads)
        case = str(key[0])
        elements = int(record.get("elements", 0))
        bounds = conditional_read_bounds(case, elements)
        if bounds is not None and not (bounds[0] <= reads <= bounds[1]):
            raise RuntimeError(f"{directory} {key}: reads {reads} outside expected bounds {bounds}")
        if case.startswith("reference-") or case.startswith("nested-") or case.startswith("database-"):
            if case.startswith("error-"):
                continue
            if int(record.get("reference_reads_p50", 0)) < 0:
                raise RuntimeError(f"{directory} {key}: invalid reference reads")
        if case.startswith("scalar-") or case.startswith("literal-") or case.startswith("array-"):
            if int(record.get("reference_reads_p50", 0)) != 0:
                raise RuntimeError(f"{directory} {key}: non-reference workload read a resolver")
    equal(f"{directory} groups", set(groups), {(case, phase) for case in expected_cases for phase in PHASES})
    for key, group in groups.items():
        equal(f"{directory} {key} sample indices", sorted(int(row["sample_index"]) for row in group), list(range(1, samples + 1)))
        if len({row.get("binary_sha256") for row in group}) != 1:
            raise RuntimeError(f"{directory} {key}: binary changed across samples")


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--results", type=Path, default=RESULTS)
    parser.add_argument("--candidate-root", type=Path, required=True)
    parser.add_argument("--candidate-freeze", type=Path, required=True)
    parser.add_argument("--expected-warmups", type=int, default=3)
    parser.add_argument("--expected-samples", type=int, default=15)
    args = parser.parse_args()
    results = args.results.resolve()
    before = load(results / "profile-inputs-before.json")
    after = load(results / "profile-inputs-after.json")
    equal("profile input hash before/after", before, after)
    equal("current profile input hash", profile_input_snapshot(), before)
    candidate_root = args.candidate_root.resolve()
    freeze = verify_candidate_freeze(candidate_root, args.candidate_freeze.resolve())
    baseline_records = rows(results / BASELINE)
    candidate_records = rows(results / CANDIDATE)
    baseline_manifest = verify_manifest(results / BASELINE, before)
    candidate_manifest = verify_manifest(results / CANDIDATE, before)
    baseline_env = load(results / BASELINE / "environment.json")
    candidate_env = load(results / CANDIDATE / "environment.json")
    equal("baseline cases", baseline_env.get("cases"), list(CONTROL_CASES))
    candidate_cases = candidate_env.get("cases")
    expected_candidate = expected_candidate_cases()
    equal("candidate case matrix", candidate_cases, list(expected_candidate))
    for key in ("rustc_verbose", "cargo", "libc", "rustflags"):
        equal(f"{key} match", baseline_env.get(key), candidate_env.get(key))
    verify_rows(results / BASELINE, baseline_records, list(CONTROL_CASES), args.expected_warmups, args.expected_samples)
    verify_rows(results / CANDIDATE, candidate_records, list(expected_candidate), args.expected_warmups, args.expected_samples)
    equal("harness lock match", baseline_manifest["before"]["harness_sha256"]["Cargo.lock"], candidate_manifest["before"]["harness_sha256"]["Cargo.lock"])
    equal("root lock match", baseline_manifest["before"]["workspace_lock_sha256"], candidate_manifest["before"]["workspace_lock_sha256"])
    summary = load(results / "capture-summary.json")
    equal("summary profile before", summary.get("profile_input_sha256_before"), before)
    equal("summary profile after", summary.get("profile_input_sha256_after"), after)
    if not summary.get("cleanup", {}).get("baseline_removed"):
        raise RuntimeError("baseline worktree cleanup receipt is false")
    print(json.dumps({"status": "ok", "baseline_records": len(baseline_records), "candidate_records": len(candidate_records), "candidate_base": freeze.get("base_commit")}, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
