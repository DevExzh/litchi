#!/usr/bin/env python3
"""Fail-closed verifier for the ODS aggregate performance receipt."""

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
AGGREGATE_NAMES = (
    "sum",
    "product",
    "sumsq",
    "sumproduct",
    "sumx2my2",
    "sumx2py2",
    "sumxmy2",
)


def expected_candidate_cases() -> tuple[str, ...]:
    # This follows the harness's all_cases() construction order.  The baseline
    # lane requests CONTROL_CASES in its stable report order; the candidate
    # lane records every available case in harness order.
    cases = [
        "scalar-control-arithmetic",
        "scalar-control-sin",
        "scalar-control-imsum",
        "database-control-dsum",
        "array-control-4x4-arithmetic",
        "array-control-4x4-sin",
        "array-control-16x16-arithmetic",
        "array-control-16x16-sin",
        "reference-scalar-arithmetic",
        "reference-array-arithmetic",
        "reference-scalar-sin",
        "reference-array-sin",
    ]
    for operation in AGGREGATE_NAMES:
        cases.append(f"scalar-aggregate-{operation}")
        for rows, columns in ((4, 4), (16, 16)):
            cases.append(f"literal-aggregate-{rows}x{columns}-{operation}")
        for rows, columns in ((16, 4), (64, 4), (256, 4), (1024, 4)):
            cases.append(f"reference-aggregate-{rows}x{columns}-{operation}")
    cases.append("literal-aggregate-4x4-sumproduct-k3")
    cases.append("reference-aggregate-1024x4-sumproduct-k3")
    for rows in (64, 256, 1024):
        cases.append(f"nested-projected-{rows}-sumproduct")
        cases.append(f"nested-if-projected-{rows}-sumproduct")
    return tuple(cases)


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


def expected_aggregate_reference_reads(case: str, rows: int, columns: int) -> int | None:
    cells = rows * columns
    if case.startswith("nested-projected-") or case.startswith("nested-if-projected-"):
        # Each projected evaluation reads the condition/outer-left/outer-right
        # lanes once.  A quadratic branch would exceed this receipt.
        return 3 * cells
    if case.startswith("reference-aggregate-"):
        if case.endswith("-sumproduct-k3"):
            return 3 * cells
        binary = case.endswith(("-sumproduct", "-sumx2my2", "-sumx2py2", "-sumxmy2"))
        return (2 if binary else 1) * cells
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
    selected = before.get("source_sha256", {})
    # The baseline owns database/numerics.rs; the candidate must own the shared
    # evaluator path after the deliberate move.  The old path must disappear.
    old_path = "crates/litchi-ods/src/codec/formula/evaluation/value/database/numerics.rs"
    new_path = "crates/litchi-ods/src/codec/formula/evaluation/numerics.rs"
    if directory.name == BASELINE:
        if not selected.get(old_path) or selected.get(new_path) is not None:
            raise RuntimeError(f"{directory}: baseline numerics ownership receipt is inconsistent")
    if directory.name == CANDIDATE:
        if selected.get(old_path) is not None or not selected.get(new_path):
            raise RuntimeError(f"{directory}: candidate numerics ownership receipt is inconsistent")
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
        expected_reads_per_repeat = int(record["reference_reads_p50"]) // int(record["repeat"])
        equal(
            f"{directory} {key} normalized reference reads",
            int(record.get("reference_reads_per_repeat", -1)),
            expected_reads_per_repeat,
        )
        expected_aggregate_reads = expected_aggregate_reference_reads(
            str(key[0]), int(record.get("rows", 0)), int(record.get("columns", 0))
        )
        if expected_aggregate_reads is not None:
            equal(
                f"{directory} {key} aggregate reference reads",
                expected_reads_per_repeat,
                expected_aggregate_reads,
            )
    equal(f"{directory} groups", set(groups), {(case, phase) for case in expected_cases for phase in PHASES})
    for key, group in groups.items():
        equal(f"{directory} {key} sample indices", sorted(int(row["sample_index"]) for row in group), list(range(1, samples + 1)))
        if len({row.get("binary_sha256") for row in group}) != 1:
            raise RuntimeError(f"{directory} {key}: binary changed across samples")
        if (
            key[0].startswith("reference-")
            or key[0].startswith("nested-projected-")
            or key[0].startswith("nested-if-projected-")
            or key[0].startswith("database-")
        ):
            if any(int(row.get("reference_reads_p50", 0)) <= 0 for row in group):
                raise RuntimeError(f"{directory} {key}: reference workload recorded no resolver reads")
        if key[0].startswith("scalar-") or key[0].startswith("literal-") or key[0].startswith("array-"):
            if any(int(row.get("reference_reads_p50", 0)) != 0 for row in group):
                raise RuntimeError(f"{directory} {key}: non-reference workload read a resolver")


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
    if candidate_cases != list(expected_candidate):
        raise RuntimeError(
            "candidate case matrix differs from the frozen harness contract:\n"
            f"expected {list(expected_candidate)!r}, observed {candidate_cases!r}"
        )
    for key in ("rustc_verbose", "cargo", "libc", "rustflags"):
        equal(f"{key} match", baseline_env.get(key), candidate_env.get(key))
    verify_rows(results / BASELINE, baseline_records, list(CONTROL_CASES), args.expected_warmups, args.expected_samples)
    verify_rows(results / CANDIDATE, candidate_records, candidate_cases, args.expected_warmups, args.expected_samples)
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
