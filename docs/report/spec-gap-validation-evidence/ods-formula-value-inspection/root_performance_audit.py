#!/usr/bin/env python3
"""Independently aggregate both complete frozen-source performance captures."""

from collections import defaultdict
import hashlib
import json
from pathlib import Path
import random
from statistics import median
import subprocess
import sys

HERE = Path(__file__).resolve().parent
REPO = HERE.parents[3]
PERFORMANCE = HERE / "performance"
sys.path.insert(0, str(PERFORMANCE))
import verify as receipts

RESULTS = PERFORMANCE / "results"
PAIRS = (
    ("earlier-complete", "diagnostic-baseline-d623f3c2ecc0c837017f700174656f0e443759a5-20260920T063738Z",
     "diagnostic-candidate-final-20260920T063808Z"),
    ("latest-complete", receipts.BASELINE, receipts.CANDIDATE),
)
ACCOUNTING = (
    "allocator_calls_p50", "requested_bytes_p50", "released_bytes_p50",
    "peak_live_delta_p50", "memory_retained_p50", "work_p50",
    "reference_reads_p50", "output_bytes_p50",
)


def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def equal_map(label, observed, expected):
    differences = [path for path in sorted(set(observed) | set(expected))
                   if observed.get(path) != expected.get(path)]
    if differences:
        raise RuntimeError(f"{label}: {len(differences)} mismatches; first: {differences[:5]}")


def groups(records):
    result = defaultdict(list)
    for record in records:
        result[(record["case"], record["phase"])].append(record)
    return result


def baseline_workspace(commit):
    """Hash the committed Rust/build closure, without a retained worktree."""
    listing = subprocess.check_output(
        ["git", "ls-tree", "-r", commit, "Cargo.toml", "crates"], cwd=REPO, text=True
    )
    blobs = {}
    for line in listing.splitlines():
        header, path = line.split("\t", 1)
        if path == "Cargo.toml" or path.endswith((".rs", "Cargo.toml", "build.rs")):
            blobs[path] = header.split()[2]
    hashes = {}
    process = subprocess.Popen(["git", "cat-file", "--batch"], cwd=REPO,
                               stdin=subprocess.PIPE, stdout=subprocess.PIPE)
    try:
        for blob in sorted(set(blobs.values())):
            process.stdin.write((blob + "\n").encode())
            process.stdin.flush()
            header = process.stdout.readline().split()
            if len(header) != 3 or header[1] != b"blob":
                raise RuntimeError(f"invalid committed blob: {blob}")
            data = process.stdout.read(int(header[2]))
            if process.stdout.read(1) != b"\n":
                raise RuntimeError("invalid git batch separator")
            hashes[blob] = hashlib.sha256(data).hexdigest()
    finally:
        process.stdin.close()
        process.stdout.close()
        if process.wait() != 0:
            raise RuntimeError("git batch hashing failed")
    result = {path: hashes[blob] for path, blob in blobs.items()}
    result["Cargo.lock"] = digest(HERE / "gates/Cargo.lock")
    return result


def interval(baseline, candidate, rng):
    shifts = sorted(
        100 * (median(rng.choices(candidate, k=len(candidate)))
               / median(rng.choices(baseline, k=len(baseline))) - 1)
        for _ in range(5000)
    )
    return [shifts[125], shifts[4874]]


def main():
    receipts.require_capture_inputs()
    profile = receipts.profile_input_snapshot()
    equal_map("profile before capture", receipts.load(RESULTS / "profile-inputs-before.json"), profile)
    equal_map("profile after capture", receipts.load(RESULTS / "profile-inputs-after.json"), profile)
    summary = receipts.load(RESULTS / "capture-summary.json")
    equal_map("summary profile before", summary["profile_input_sha256_before"], profile)
    equal_map("summary profile after", summary["profile_input_sha256_after"], profile)
    receipts.equal("baseline worktree cleanup", summary["cleanup"]["baseline_removed"], True)
    receipts.verify_preflight(RESULTS / "preflight-before-timing", list(receipts.expected_candidate_cases()))
    freeze = receipts.load(HERE / "gates/freeze.json")
    gate = receipts.load(HERE / "gates/source-before.json")
    committed = baseline_workspace(freeze["base_commit"])
    expected_candidate = {
        path: value for path, value in gate["workspace_source_sha256"].items()
        if path in {"Cargo.toml", "Cargo.lock"}
        or (path.startswith("crates/") and path.endswith((".rs", "Cargo.toml", "build.rs")))
    }
    captures = []
    candidates = []
    baselines = []
    rng = random.Random(20260920)
    for label, baseline_name, candidate_name in PAIRS:
        records = []
        manifests = []
        for name, cases in ((baseline_name, receipts.MATCHED_CONTROL_CASES),
                            (candidate_name, receipts.expected_candidate_cases())):
            directory = RESULTS / name
            data = receipts.rows(directory)
            manifests.append(receipts.verify_manifest(directory, profile))
            receipts.verify_rows(directory, data, list(cases), 3, 15)
            receipts.verify_preflight(directory, list(cases))
            records.append(data)
        baselines.append(manifests[0]["before"])
        candidates.append(manifests[1])
        equal_map(f"{label} committed baseline closure", manifests[0]["before"]["workspace_source_sha256"], committed)
        equal_map(f"{label} candidate gate closure", manifests[1]["before"]["workspace_source_sha256"], expected_candidate)
        equal_map(f"{label} candidate selected sources", manifests[1]["before"]["source_sha256"], freeze["selected_files"])
        receipts.equal(f"{label} matched harness", manifests[0]["before"]["harness_sha256"], manifests[1]["before"]["harness_sha256"])
        environments = [receipts.load(RESULTS / name / "environment.json") for name in (baseline_name, candidate_name)]
        for manifest in manifests:
            receipts.equal(f"{label} baseline checkout identity", manifest["before"]["git_head"], freeze["base_commit"])
            receipts.equal(f"{label} stable checkout identity", manifest["after"]["git_head"], freeze["base_commit"])
        receipts.equal(f"{label} baseline cases", environments[0]["cases"], list(receipts.MATCHED_CONTROL_CASES))
        receipts.equal(f"{label} candidate cases", environments[1]["cases"], list(receipts.expected_candidate_cases()))
        for key in ("rustc_verbose", "cargo", "libc", "rustflags", "contract_sha256"):
            receipts.equal(f"{label} matched {key}", environments[0][key], environments[1][key])
        receipts.equal(f"{label} contract pin", environments[1]["contract_sha256"], receipts.CONTRACT_SHA256)
        baseline, candidate = map(groups, records)
        comparisons = []
        for key in sorted(baseline):
            b, c = baseline[key], candidate[key]
            changed = [metric for metric in ACCOUNTING
                       if {row[metric] for row in b} != {row[metric] for row in c}]
            if changed:
                raise RuntimeError(f"{label} {key}: accounting changed: {changed}")
            bt = [row["elapsed_ns_p50"] / row["repeat"] for row in b]
            ct = [row["elapsed_ns_p50"] / row["repeat"] for row in c]
            br, cr = (median(row["rss_kib"] for row in rows) for rows in (b, c))
            latency = 100 * (median(ct) / median(bt) - 1)
            rss = 100 * (cr / br - 1)
            comparisons.append({
                "case": key[0], "phase": key[1],
                "baseline_ns_per_repeat": median(bt), "candidate_ns_per_repeat": median(ct),
                "latency_delta_percent": latency,
                "latency_bootstrap_95_percent": interval(bt, ct, rng),
                "rss_delta_percent": rss, "rss_delta_kib": cr - br,
                "review_trigger": latency > 5 or rss > 5,
            })
        captures.append({
            "label": label, "baseline": baseline_name, "candidate": candidate_name,
            "baseline_samples": len(records[0]), "candidate_samples": len(records[1]),
            "matched_groups": len(comparisons), "accounting_unchanged": list(ACCOUNTING),
            "manifest_sha256": {name: digest(RESULTS / name / "source-manifest.json")
                                for name in (baseline_name, candidate_name)},
            "comparisons": comparisons,
        })
    for key in ("source_sha256", "workspace_source_sha256", "profile_input_sha256", "harness_sha256"):
        receipts.equal(f"candidate {key} across captures", candidates[0]["before"][key], candidates[1]["before"][key])
        receipts.equal(f"baseline {key} across captures", baselines[0][key], baselines[1][key])
    receipts.equal("candidate binary across captures", candidates[0]["binary_sha256"], candidates[1]["binary_sha256"])
    result = {
        "status": "verified; performance review triggers retained",
        "verified_complete_samples": sum(c["baseline_samples"] + c["candidate_samples"] for c in captures),
        "bootstrap": {"seed": 20260920, "resamples": 5000, "method": "independent unpaired median ratios; percentile interval"},
        "captures": captures,
    }
    (HERE / "root-performance-audit.json").write_text(json.dumps(result, indent=2) + "\n")
    print(json.dumps({"status": "ok", "samples": result["verified_complete_samples"],
                      "review_triggers": {c["label"]: sum(row["review_trigger"] for row in c["comparisons"]) for c in captures}}))


if __name__ == "__main__":
    main()
