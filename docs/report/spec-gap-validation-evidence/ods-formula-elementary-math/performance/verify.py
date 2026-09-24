#!/usr/bin/env python3
"""Fail-closed validation for the bounded ODS elementary-math profile."""

from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any

from run_profile import BASELINE_COMMIT, CONTROL_CASES, PHASES, digest, verify_candidate_freeze


HERE = Path(__file__).resolve().parent
RESULTS = HERE / "results"
BASELINE = "baseline-8ef0057e5"
CANDIDATE = "candidate-final"


def load_json(path: Path) -> dict[str, Any]:
    value = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(value, dict):
        raise RuntimeError(f"expected object in {path}")
    return value


def load_rows(directory: Path) -> list[dict[str, Any]]:
    path = directory / "measurements.jsonl"
    if not path.is_file():
        raise RuntimeError(f"missing measurements: {path}")
    rows = [json.loads(line) for line in path.read_text(encoding="utf-8").splitlines() if line.strip()]
    if not rows:
        raise RuntimeError(f"empty measurements: {path}")
    return rows


def assert_equal(label: str, actual: Any, expected: Any) -> None:
    if actual != expected:
        raise RuntimeError(f"{label}: expected {expected!r}, observed {actual!r}")


def verify_manifest(directory: Path, profile_hashes: dict[str, str]) -> dict[str, Any]:
    manifest = load_json(directory / "source-manifest.json")
    before = manifest.get("before")
    after = manifest.get("after")
    if not isinstance(before, dict) or not isinstance(after, dict):
        raise RuntimeError(f"{directory}: malformed source manifest")
    for key in ("source_sha256", "workspace_source_sha256", "profile_input_sha256"):
        assert_equal(f"{directory} {key} stable", before.get(key), after.get(key))
    assert_equal(f"{directory} profile inputs", before.get("profile_input_sha256"), profile_hashes)
    assert_equal(f"{directory} selected source flag", manifest.get("source_sha256_unchanged"), True)
    assert_equal(
        f"{directory} workspace source flag",
        manifest.get("workspace_source_sha256_unchanged"),
        True,
    )
    assert_equal(f"{directory} profile source flag", manifest.get("profile_input_sha256_unchanged"), True)
    assert_equal(f"{directory} Git HEAD stable", manifest.get("git_head_unchanged"), True)
    lock = before.get("harness_sha256", {}).get("Cargo.lock")
    if not lock:
        raise RuntimeError(f"{directory}: missing harness Cargo.lock hash")
    if not before.get("fixture_sha256"):
        raise RuntimeError(f"{directory}: missing fixture hash")
    cleanup = load_json(directory / "target-cleanup.json")
    assert_equal(f"{directory} Cargo target cleanup", cleanup.get("removed"), True)
    return manifest


def verify_environment(
    directory: Path,
    *,
    expected_cases: list[str],
    expected_warmups: int,
    expected_samples: int,
    expected_scope: str,
) -> dict[str, Any]:
    environment = load_json(directory / "environment.json")
    assert_equal(f"{directory} cases", environment.get("cases"), expected_cases)
    assert_equal(f"{directory} phases", environment.get("phases"), list(PHASES))
    assert_equal(f"{directory} warmups", environment.get("warmups_per_child"), expected_warmups)
    assert_equal(f"{directory} samples", environment.get("samples_per_group"), expected_samples)
    assert_equal(f"{directory} scope", environment.get("case_scope"), expected_scope)
    for key in ("rustc", "rustc_verbose", "cargo", "libc", "cpu", "binary_sha256", "profile_input_sha256"):
        if not environment.get(key):
            raise RuntimeError(f"{directory}: missing environment field {key}")
    return environment


def verify_rows(
    directory: Path,
    rows: list[dict[str, Any]],
    *,
    expected_cases: list[str],
    expected_samples: int,
    require_supported: bool,
) -> None:
    expected_records = len(expected_cases) * len(PHASES) * expected_samples
    assert_equal(f"{directory} record count", len(rows), expected_records)
    groups: dict[tuple[str, str], list[dict[str, Any]]] = {}
    for row in rows:
        key = (row.get("case"), row.get("phase"))
        if key[0] not in expected_cases or key[1] not in PHASES:
            raise RuntimeError(f"{directory}: unexpected group {key}")
        groups.setdefault(key, []).append(row)
        if require_supported and row.get("supported") is not True:
            raise RuntimeError(f"{directory} {key}: unsupported result")
        if row.get("validation_scope", "").startswith("one untimed direct f64") is False:
            raise RuntimeError(f"{directory} {key}: oracle scope missing or overstated")
        if int(row.get("rss_kib", 0)) <= 0 or int(row.get("elapsed_ns_p50", 0)) <= 0:
            raise RuntimeError(f"{directory} {key}: invalid elapsed/RSS receipt")
        if int(row.get("repeat", 0)) <= 0:
            raise RuntimeError(f"{directory} {key}: invalid repeat")
        if row.get("raw_stderr") is not None:
            stderr = directory / row["raw_stderr"]
            if not stderr.is_file() or stderr.read_text(encoding="utf-8") == "":
                raise RuntimeError(f"{directory} {key}: empty retained stderr")
        for raw_key in ("raw_stdout", "raw_time"):
            if not (directory / row[raw_key]).is_file():
                raise RuntimeError(f"{directory} {key}: missing {raw_key}")
        samples = row.get("samples")
        if not isinstance(samples, list) or len(samples) != 1:
            raise RuntimeError(f"{directory} {key}: each child must emit one measured sample")
        sample = samples[0]
        for metric in (
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
        ):
            if int(sample.get(metric, -1)) < 0:
                raise RuntimeError(f"{directory} {key}: negative {metric}")
        if sample["live_before"] != sample["live_after"]:
            raise RuntimeError(f"{directory} {key}: live allocator bytes do not balance")
        if int(sample["requested_bytes"]) < int(sample["released_bytes"]):
            raise RuntimeError(f"{directory} {key}: released bytes exceed requested bytes")
    expected_keys = {(case, phase) for case in expected_cases for phase in PHASES}
    assert_equal(f"{directory} groups", set(groups), expected_keys)
    for key, group in groups.items():
        indices = sorted(int(row["sample_index"]) for row in group)
        assert_equal(f"{directory} {key} sample indices", indices, list(range(1, expected_samples + 1)))
        binaries = {row.get("binary_sha256") for row in group}
        if len(binaries) != 1:
            raise RuntimeError(f"{directory} {key}: binary changed across samples")


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--results", type=Path, default=RESULTS)
    parser.add_argument("--candidate-root", type=Path, required=True)
    parser.add_argument("--candidate-freeze", type=Path, required=True)
    parser.add_argument("--expected-warmups", type=int, default=3)
    parser.add_argument("--expected-samples", type=int, default=15)
    args = parser.parse_args()
    if args.expected_warmups < 0 or args.expected_samples <= 0:
        parser.error("expected warmups must be nonnegative and samples must be positive")
    results = args.results.resolve()
    profile_before = load_json(results / "profile-inputs-before.json")
    profile_after = load_json(results / "profile-inputs-after.json")
    assert_equal("profile input hashes before/after", profile_before, profile_after)
    freeze = verify_candidate_freeze(args.candidate_root.resolve(), args.candidate_freeze.resolve())
    base_commit = str(freeze.get("base_commit", ""))
    if not (base_commit == BASELINE_COMMIT or base_commit.startswith(BASELINE_COMMIT)):
        raise RuntimeError("candidate freeze does not name the required baseline")

    baseline_dir = results / BASELINE
    candidate_dir = results / CANDIDATE
    baseline_manifest = verify_manifest(baseline_dir, profile_before)
    candidate_manifest = verify_manifest(candidate_dir, profile_before)
    baseline_environment = verify_environment(
        baseline_dir,
        expected_cases=list(CONTROL_CASES),
        expected_warmups=args.expected_warmups,
        expected_samples=args.expected_samples,
        expected_scope="matched-controls",
    )
    candidate_environment = verify_environment(
        candidate_dir,
        expected_cases=load_json(candidate_dir / "environment.json")["cases"],
        expected_warmups=args.expected_warmups,
        expected_samples=args.expected_samples,
        expected_scope="all-named-cases",
    )
    candidate_cases = candidate_environment["cases"]
    if len(candidate_cases) != len(set(candidate_cases)) or len(candidate_cases) < 35:
        raise RuntimeError("candidate does not contain the full named workload matrix")
    if set(CONTROL_CASES) - set(candidate_cases):
        raise RuntimeError("candidate is missing matched control cases")
    verify_rows(
        baseline_dir,
        load_rows(baseline_dir),
        expected_cases=list(CONTROL_CASES),
        expected_samples=args.expected_samples,
        require_supported=True,
    )
    verify_rows(
        candidate_dir,
        load_rows(candidate_dir),
        expected_cases=candidate_cases,
        expected_samples=args.expected_samples,
        require_supported=True,
    )
    assert_equal("harness lock match", baseline_manifest["before"]["harness_sha256"]["Cargo.lock"], candidate_manifest["before"]["harness_sha256"]["Cargo.lock"])
    assert_equal("fixture match", baseline_manifest["before"]["fixture_sha256"], candidate_manifest["before"]["fixture_sha256"])
    for key in ("rustc_verbose", "cargo", "libc", "rustflags"):
        assert_equal(f"{key} match", baseline_environment[key], candidate_environment[key])
    summary = load_json(results / "capture-summary.json")
    assert_equal("summary profile before", summary.get("profile_input_sha256_before"), profile_before)
    assert_equal("summary profile after", summary.get("profile_input_sha256_after"), profile_after)
    if not summary.get("cleanup", {}).get("baseline_removed"):
        raise RuntimeError("baseline worktree cleanup receipt is false")
    print(json.dumps({"status": "ok", "baseline_records": len(load_rows(baseline_dir)), "candidate_records": len(load_rows(candidate_dir))}, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
