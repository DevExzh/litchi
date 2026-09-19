#!/usr/bin/env python3
"""Fail-closed verifier for descriptive-statistics performance receipts."""

from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any

from run_profile import (
    BASELINE_COMMIT,
    CONTRACT_SHA256,
    MATCHED_CONTROL_CASES,
    PHASES,
    profile_input_snapshot,
    require_capture_inputs,
    verify_candidate_freeze,
)

HERE = Path(__file__).resolve().parent
RESULTS = HERE / "results"
BASELINE = f"baseline-{BASELINE_COMMIT}"
CANDIDATE = "candidate-final"
FUNCTIONS = ("avedev", "devsq", "geomean", "harmean", "kurt", "skew", "skewp")
LANES = (
    "scalar",
    "inline",
    "reference-descriptive-64",
    "reference-descriptive-256",
    "reference-descriptive-1024",
    "projected-descriptive-64",
    "projected-descriptive-256",
    "projected-descriptive-1024",
    "domain",
    "list",
    "error",
    "cancellation",
    "resource",
)


def expected_candidate_cases() -> tuple[str, ...]:
    cases = list(MATCHED_CONTROL_CASES)
    for function in FUNCTIONS:
        for lane in LANES:
            if lane == "list":
                lane = "list-refusal" if function in {"devsq", "skewp"} else "list-admit"
            elif lane.startswith(("reference-descriptive-", "projected-descriptive-")):
                cases.append(f"{lane}-{function}")
                continue
            cases.append(f"{lane}-descriptive-{function}")
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
    if case in {"reference-control-average", "reference-control-counta"}:
        return elements, elements
    for size in (64, 256, 1024):
        if case.startswith(f"reference-descriptive-{size}-"):
            function = case.rsplit("-", 1)[-1]
            passes = 2 if function == "avedev" else 1
            expected = size * 4 * passes
            return expected, expected
        if case.startswith(f"projected-descriptive-{size}-"):
            function = case.rsplit("-", 1)[-1]
            passes = 2 if function == "avedev" else 1
            expected = size * 4 * passes
            return expected, expected
    if case.startswith(("scalar-descriptive-", "inline-descriptive-", "domain-descriptive-")):
        return 0, 0
    if case.startswith("list-refusal-descriptive-"):
        return 0, 0
    if case.startswith("list-admit-descriptive-"):
        function = case.rsplit("-", 1)[-1]
        passes = 2 if function == "avedev" else 1
        expected = 16 * passes
        return expected, expected
    if case.startswith("resource-descriptive-"):
        return 0, 0
    if case.startswith("cancellation-descriptive-"):
        return 1, 1
    if case.startswith("error-descriptive-"):
        # A retained source formula error is final after the complete primary
        # scan, so a centered reducer skips an otherwise meaningless replay.
        return elements, elements
    # Future adaptive harmonic-replay rows may carry an owner-specific input
    # profile; ordinary rows above remain exact and fail closed here.
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
        ):
            if int(sample.get(field, -1)) < 0:
                raise RuntimeError(f"{directory} {key}: negative {field}")
        if sample["live_before"] != sample["live_after"]:
            raise RuntimeError(f"{directory} {key}: allocator bytes do not balance")
        if int(sample["released_bytes"]) > int(sample["requested_bytes"]):
            raise RuntimeError(f"{directory} {key}: released bytes exceed requested")
        repeat = int(record["repeat"])
        reads = int(record.get("reference_reads_p50", 0)) // repeat
        equal(f"{directory} {key} normalized reference reads", int(record.get("reference_reads_per_repeat", -1)), reads)
        bounds = direct_read_bound(str(key[0]), int(record.get("elements", 0)))
        if bounds is not None and not bounds[0] <= reads <= bounds[1]:
            raise RuntimeError(f"{directory} {key}: reads {reads} outside expected bounds {bounds}")
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
