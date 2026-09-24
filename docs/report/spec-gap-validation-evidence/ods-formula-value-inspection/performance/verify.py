#!/usr/bin/env python3
"""Fail-closed verifier for value-inspection performance receipts."""

from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any

from run_profile import (
    BASELINE_COMMIT,
    CONTRACT_SHA256,
    MATCHED_CONTROL_CASES,
    INSPECTION_CASES,
    PHASES,
    preflight_read_bound,
    profile_input_snapshot,
    require_capture_inputs,
    verify_candidate_freeze,
)

HERE = Path(__file__).resolve().parent
RESULTS = HERE / "results"
BASELINE = f"baseline-{BASELINE_COMMIT}"
CANDIDATE = "candidate-final"
def expected_candidate_cases() -> tuple[str, ...]:
    cases = list(MATCHED_CONTROL_CASES)
    cases.extend(INSPECTION_CASES)
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


def direct_read_bound(case: str, elements: int) -> tuple[int, int] | None:
    if case == "reference-aggregate-64x4-sum":
        return elements, elements
    if case == "reference-conditional-256x4-sumifs":
        return 7 * elements // 4, 7 * elements // 4
    if case in {"database-control-dsum", "database-control-dvar", "database-control-dstdev"}:
        return 7, 7
    if case in {"reference-control-average", "reference-control-counta"}:
        return elements, elements
    if case == "n-reference-intersection":
        return 1, 1
    if case == "type-reference-scan" or case.startswith("reference-inspection-"):
        return elements, elements
    if case.startswith("lazy-if-cache-"):
        return 2, 2
    if case.startswith(("shape-refusal-", "resource-inspection-")) or case == "numbervalue-invalid-separator":
        return 0, 0
    # Cancellation is verified as one total read per child below; dividing the
    # total by four internal repeats intentionally produces integer zero.
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
        if not record.get("raw_stdout") or not (directory / record["raw_stdout"]).is_file():
            raise RuntimeError(f"{directory} {key}: raw stdout missing")
        if not (directory / record["raw_time"]).is_file():
            raise RuntimeError(f"{directory} {key}: raw time missing")
        raw_samples = record.get("samples")
        if not isinstance(raw_samples, list) or len(raw_samples) != 1:
            raise RuntimeError(f"{directory} {key}: child did not emit one sample")
        sample = raw_samples[0]
        for field in (
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
        ):
            if int(sample.get(field, -1)) < 0:
                raise RuntimeError(f"{directory} {key}: negative {field}")
        if sample["live_before"] != sample["live_after"]:
            raise RuntimeError(f"{directory} {key}: allocator bytes do not balance")
        if int(sample["released_bytes"]) > int(sample["requested_bytes"]):
            raise RuntimeError(f"{directory} {key}: released bytes exceed requested")
        repeat = int(record["repeat"])
        total_reads = int(record.get("reference_reads_p50", 0))
        reads = total_reads // repeat
        equal(f"{directory} {key} normalized reference reads", int(record.get("reference_reads_per_repeat", -1)), reads)
        input_bytes = int(record.get("input_bytes", -1))
        output_bytes = int(record.get("output_bytes_p50", -1))
        bytes_per_repeat = int(record.get("bytes_per_repeat_p50", -1))
        if input_bytes < 0 or output_bytes < 0:
            raise RuntimeError(f"{directory} {key}: invalid byte metrics")
        equal(
            f"{directory} {key} normalized byte throughput",
            bytes_per_repeat,
            input_bytes + output_bytes // repeat,
        )
        if str(key[0]).startswith("cancellation-inspection-"):
            equal(f"{directory} {key} cancellation total reads", total_reads, 1)
            equal(f"{directory} {key} cancellation repeat count", repeat, 4)
        bounds = direct_read_bound(str(key[0]), int(record.get("elements", 0)))
        if bounds is not None and not bounds[0] <= reads <= bounds[1]:
            raise RuntimeError(f"{directory} {key}: reads {reads} outside expected bounds {bounds}")
    equal(f"{directory} groups", set(groups), {(case, phase) for case in expected_cases for phase in PHASES})
    for key, group in groups.items():
        equal(f"{directory} {key} sample indices", sorted(int(row["sample_index"]) for row in group), list(range(1, samples + 1)))
        if len({row.get("binary_sha256") for row in group}) != 1:
            raise RuntimeError(f"{directory} {key}: binary changed across samples")


def verify_preflight(directory: Path, expected_cases: list[str]) -> None:
    preflight = load(directory / "preflight.json")
    equal(f"{directory} preflight status", preflight.get("status"), "ok")
    stdout = directory / str(preflight.get("stdout", ""))
    if not stdout.is_file():
        raise RuntimeError(f"{directory}: preflight stdout missing")
    reads = preflight.get("reference_reads")
    if not isinstance(reads, dict):
        raise RuntimeError(f"{directory}: preflight reference reads missing")
    equal(f"{directory} preflight cases", sorted(reads), sorted(expected_cases))
    for case in expected_cases:
        equal(
            f"{directory} preflight reads {case}",
            int(reads[case]),
            preflight_read_bound(case),
        )


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--results", type=Path, default=RESULTS)
    parser.add_argument("--candidate-root", type=Path, required=True)
    parser.add_argument("--candidate-freeze", type=Path, required=True)
    parser.add_argument("--expected-warmups", type=int, default=3)
    parser.add_argument("--expected-samples", type=int, default=15)
    args = parser.parse_args()
    require_capture_inputs()
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
    equal("baseline cases", baseline_env.get("cases"), list(MATCHED_CONTROL_CASES))
    expected_candidate = expected_candidate_cases()
    equal("candidate case matrix", candidate_env.get("cases"), list(expected_candidate))
    for key in ("rustc_verbose", "cargo", "libc", "rustflags"):
        equal(f"{key} match", baseline_env.get(key), candidate_env.get(key))
    equal("baseline contract hash", baseline_env.get("contract_sha256"), CONTRACT_SHA256)
    equal("candidate contract hash", candidate_env.get("contract_sha256"), CONTRACT_SHA256)
    verify_rows(results / BASELINE, baseline_records, list(MATCHED_CONTROL_CASES), args.expected_warmups, args.expected_samples)
    verify_rows(results / CANDIDATE, candidate_records, list(expected_candidate), args.expected_warmups, args.expected_samples)
    verify_preflight(results / BASELINE, list(MATCHED_CONTROL_CASES))
    verify_preflight(results / CANDIDATE, list(expected_candidate))
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
