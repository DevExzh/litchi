#!/usr/bin/env python3
"""Audit the retained reference-metadata performance pair.

The capture runner removes its detached checkout and build targets before the
results are accepted.  This audit therefore reconstructs the committed
baseline workspace from Git and checks the candidate workspace against the
retained gate manifests.  It consumes raw receipts only; it never starts a
benchmark or requires the temporary candidate checkout to survive.

The case list, sample count, accounting fields, and capture identity come from
the current profile and retained receipts.  No historical diagnostic capture
name or outcome total is embedded here.
"""

from __future__ import annotations

from collections import defaultdict
import hashlib
import json
import fnmatch
from pathlib import Path
import random
from statistics import median
import subprocess
import sys
from typing import Any


HERE = Path(__file__).resolve().parent
REPO = HERE.parents[3]
PERFORMANCE = HERE / "performance"
RESULTS = PERFORMANCE / "results"
GATES = HERE / "gates"

sys.path.insert(0, str(PERFORMANCE))
import run_profile as profile  # noqa: E402
import verify as receipts  # noqa: E402


class AuditError(RuntimeError):
    """A retained performance receipt is incomplete or inconsistent."""


def load(path: Path) -> Any:
    if not path.is_file():
        raise AuditError(f"missing retained performance input: {path}")
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeDecodeError, json.JSONDecodeError) as error:
        raise AuditError(f"invalid JSON retained performance input {path}: {error}") from error


def digest(path: Path) -> str:
    hasher = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            hasher.update(chunk)
    return hasher.hexdigest()


def equal(label: str, observed: Any, expected: Any) -> None:
    if observed != expected:
        raise AuditError(f"{label}: expected {expected!r}, observed {observed!r}")


def equal_map(label: str, observed: dict[str, Any], expected: dict[str, Any]) -> None:
    if not isinstance(observed, dict) or not isinstance(expected, dict):
        raise AuditError(f"{label}: source maps are malformed")
    differences = [
        path
        for path in sorted(set(observed) | set(expected))
        if observed.get(path) != expected.get(path)
    ]
    if differences:
        raise AuditError(f"{label}: {len(differences)} mismatches; first {differences[:5]}")


def safe_relative(directory: Path, value: Any, label: str) -> Path:
    if not isinstance(value, str) or not value or Path(value).is_absolute():
        raise AuditError(f"{label} is not a relative path: {value!r}")
    path = (directory / value).resolve()
    try:
        path.relative_to(directory.resolve())
    except ValueError as error:
        raise AuditError(f"{label} escapes its capture directory: {value!r}") from error
    return path


def git_blob_hashes(commit: str) -> dict[str, str]:
    """Return SHA-256 hashes for the committed Rust/build workspace closure."""

    listing = subprocess.check_output(
        ["git", "ls-tree", "-r", commit, "Cargo.toml", "crates"],
        cwd=REPO,
        text=True,
    )
    blobs: dict[str, str] = {}
    for line in listing.splitlines():
        header, relative = line.split("\t", 1)
        if relative == "Cargo.toml" or relative.endswith((".rs", "Cargo.toml", "build.rs")):
            fields = header.split()
            if len(fields) != 3 or fields[1] != "blob":
                raise AuditError(f"invalid Git tree entry: {line!r}")
            blobs[relative] = fields[2]
    if not blobs:
        raise AuditError(f"Git baseline {commit} has no workspace source closure")
    result: dict[str, str] = {}
    process = subprocess.Popen(
        ["git", "cat-file", "--batch"],
        cwd=REPO,
        stdin=subprocess.PIPE,
        stdout=subprocess.PIPE,
    )
    assert process.stdin is not None and process.stdout is not None
    try:
        for oid in sorted(set(blobs.values())):
            process.stdin.write((oid + "\n").encode())
            process.stdin.flush()
            header = process.stdout.readline().split()
            if len(header) != 3 or header[1] != b"blob":
                raise AuditError(f"invalid committed blob response for {oid}")
            data = process.stdout.read(int(header[2]))
            if process.stdout.read(1) != b"\n":
                raise AuditError(f"invalid Git batch separator for {oid}")
            result[oid] = hashlib.sha256(data).hexdigest()
    finally:
        process.stdin.close()
        process.stdout.close()
        if process.wait() != 0:
            raise AuditError("git cat-file --batch failed")
    return {relative: result[oid] for relative, oid in blobs.items()}


def baseline_workspace(commit: str, gate_lock: Path) -> dict[str, str]:
    result = git_blob_hashes(commit)
    result["Cargo.lock"] = digest(gate_lock)
    return result


def committed_digest(commit: str, relative: str) -> str | None:
    """Return the committed SHA-256 for a path, if the baseline contains it."""

    try:
        content = subprocess.check_output(
            ["git", "show", f"{commit}:{relative}"],
            cwd=REPO,
            stderr=subprocess.DEVNULL,
        )
    except subprocess.CalledProcessError:
        return None
    return digest_bytes(content)


def baseline_profile_paths(commit: str) -> set[str]:
    """Reconstruct the profile source path set without a baseline checkout."""

    listing = subprocess.check_output(
        ["git", "ls-tree", "-r", "--name-only", commit],
        cwd=REPO,
        text=True,
    ).splitlines()
    paths = set(profile.SOURCE_FILES)
    for pattern in profile.SOURCE_FILE_GLOBS:
        paths.update(path for path in listing if fnmatch.fnmatch(path, pattern))
    return paths


def rows(directory: Path) -> list[dict[str, Any]]:
    path = directory / "measurements.jsonl"
    if not path.is_file():
        raise AuditError(f"missing measurements: {path}")
    result: list[dict[str, Any]] = []
    for line_number, line in enumerate(path.read_text(encoding="utf-8").splitlines(), 1):
        if not line.strip():
            continue
        try:
            value = json.loads(line)
        except json.JSONDecodeError as error:
            raise AuditError(f"invalid measurement JSON at {path}:{line_number}: {error}") from error
        if not isinstance(value, dict):
            raise AuditError(f"measurement at {path}:{line_number} is not an object")
        result.append(value)
    if not result:
        raise AuditError(f"measurements are empty: {path}")
    return result


def grouped(records: list[dict[str, Any]]) -> dict[tuple[str, str], list[dict[str, Any]]]:
    result: dict[tuple[str, str], list[dict[str, Any]]] = defaultdict(list)
    for record in records:
        case = record.get("case")
        phase = record.get("phase")
        if not isinstance(case, str) or not isinstance(phase, str):
            raise AuditError("measurement is missing case or phase")
        result[(case, phase)].append(record)
    return dict(result)


def verify_raw_paths(directory: Path, records: list[dict[str, Any]]) -> None:
    for record in records:
        for field in ("raw_stdout", "raw_stderr", "raw_time"):
            value = record.get(field)
            if value is None:
                if field == "raw_stderr":
                    continue
                raise AuditError(f"{directory}: measurement lacks {field}")
            path = safe_relative(directory, value, f"{directory} {field}")
            if not path.is_file():
                raise AuditError(f"{directory}: missing raw receipt {value}")


def expected_cases_for_capture(capture: dict[str, Any]) -> list[str]:
    scope = capture.get("case_scope")
    if scope == "matched-controls":
        return list(profile.MATCHED_CONTROL_CASES)
    if scope == "all-named-cases":
        return list(receipts.expected_candidate_cases())
    raise AuditError(f"unknown capture case scope: {scope!r}")


def verify_capture(
    capture: dict[str, Any],
    *,
    profile_hashes: dict[str, str],
    warmups: int,
    samples: int,
) -> tuple[list[dict[str, Any]], dict[str, Any], dict[str, Any]]:
    label = capture.get("label")
    if not isinstance(label, str) or not label:
        raise AuditError("capture has no label")
    directory = RESULTS / label
    if not directory.is_dir():
        raise AuditError(f"capture directory is absent: {directory}")
    expected_cases = expected_cases_for_capture(capture)
    equal(f"{label} capture cases", capture.get("cases"), expected_cases)
    records = receipts.rows(directory)
    equal(f"{label} capture record count", capture.get("records"), len(records))
    manifest = receipts.verify_manifest(directory, profile_hashes)
    environment = load(directory / "environment.json")
    equal(f"{label} environment cases", environment.get("cases"), expected_cases)
    equal(f"{label} environment phases", environment.get("phases"), list(profile.PHASES))
    if environment.get("available_cases") is not None:
        available = environment.get("available_cases")
        if not isinstance(available, list) or not set(expected_cases).issubset(available):
            raise AuditError(f"{label} environment case availability is incomplete")
    receipts.verify_rows(directory, records, expected_cases, warmups, samples)
    receipts.verify_preflight(directory, expected_cases)
    verify_raw_paths(directory, records)
    equal(f"{label} source git head", manifest["before"].get("git_head"), profile.BASELINE_COMMIT)
    equal(f"{label} stable source git head", manifest["after"].get("git_head"), profile.BASELINE_COMMIT)
    equal(f"{label} capture source git head", capture.get("source_git_head"), profile.BASELINE_COMMIT)
    equal(f"{label} capture binary", capture.get("binary_sha256"), manifest.get("binary_sha256"))
    return records, manifest, environment


def compiled_workspace_map(workspace: dict[str, Any]) -> dict[str, Any]:
    """Keep the compiled Rust/build closure shared by gate and profiler.

    The gate manifest additionally records literal ``include_*`` payloads
    (fixtures and evidence blobs).  The performance source snapshot's
    ``workspace_source_files`` intentionally records Cargo manifests and Rust/
    build files only, so those payloads are checked by gate source custody but
    are not expected in a retained profiler manifest.
    """

    return {
        path: value
        for path, value in workspace.items()
        if path in {"Cargo.toml", "Cargo.lock"}
        or (path.startswith("crates/") and path.endswith((".rs", "Cargo.toml", "build.rs")))
    }


def verify_candidate_source(
    manifest: dict[str, Any],
    freeze: dict[str, Any],
    gate_before: dict[str, Any],
    gate_lock: Path,
) -> list[str]:
    selected = freeze.get("selected_files")
    if not isinstance(selected, dict) or not selected:
        raise AuditError("performance freeze has no selected source map")
    source_map = manifest.get("before", {}).get("source_sha256")
    equal_map("candidate selected source closure", source_map, selected)
    for relative, expected in selected.items():
        path = gate_lock if relative == "Cargo.lock" else REPO / relative
        if not path.is_file():
            raise AuditError(f"candidate selected source is absent after checkout cleanup: {relative}")
        equal(f"candidate selected source {relative}", digest(path), expected)
    gate_workspace = gate_before.get("workspace_source_sha256")
    candidate_workspace = manifest.get("before", {}).get("workspace_source_sha256")
    if not isinstance(gate_workspace, dict) or not isinstance(candidate_workspace, dict):
        raise AuditError("candidate/gate workspace closure is malformed")
    expected_compiled = compiled_workspace_map(gate_workspace)
    equal_map("candidate gate compiled workspace closure", candidate_workspace, expected_compiled)
    equal("candidate workspace lock", manifest["before"].get("workspace_lock_sha256"), digest(gate_lock))
    return sorted(set(gate_workspace) - set(expected_compiled))


def verify_gate_payloads(
    gate_before: dict[str, Any], freeze: dict[str, Any], excluded: list[str]
) -> None:
    """Verify gate-only include payloads against the committed baseline.

    The profiler records the compiled Cargo/Rust closure.  The gate manifest
    also records literal include dependencies, which are intentionally absent
    from that profiler closure.  They remain part of source custody and must
    therefore be checked independently against the baseline Git object.
    """

    workspace = gate_before.get("workspace_source_sha256")
    if not isinstance(workspace, dict):
        raise AuditError("gate workspace closure is malformed")
    selected = freeze.get("selected_files")
    if not isinstance(selected, dict):
        raise AuditError("performance freeze selected source map is malformed")
    for relative in excluded:
        observed = workspace.get(relative)
        # Newly authored selected evidence can be a literal include payload
        # even though it is not part of the profiler's compiled closure.
        # Its custody is the frozen selected hash; all other payloads must be
        # unchanged baseline Git objects.
        expected = selected.get(relative)
        if expected is None:
            expected = committed_digest(profile.BASELINE_COMMIT, relative)
        if expected is None:
            raise AuditError(f"gate-only workspace payload is absent from baseline Git: {relative}")
        equal(f"gate-only workspace payload {relative}", observed, expected)


def verify_baseline_source(manifest: dict[str, Any], gate_lock: Path) -> None:
    source_map = manifest.get("before", {}).get("source_sha256")
    if not isinstance(source_map, dict):
        raise AuditError("baseline selected source closure is malformed")
    equal(
        "baseline selected source path set",
        set(source_map),
        baseline_profile_paths(profile.BASELINE_COMMIT),
    )
    for relative, observed in source_map.items():
        expected = digest(gate_lock) if relative == "Cargo.lock" else None
        if relative != "Cargo.lock":
            try:
                data = subprocess.check_output(
                    ["git", "show", f"{profile.BASELINE_COMMIT}:{relative}"],
                    cwd=REPO,
                    stderr=subprocess.DEVNULL,
                )
            except subprocess.CalledProcessError:
                data = None
            expected = digest_bytes(data) if data is not None else None
        equal(f"baseline selected source {relative}", observed, expected)
    equal("baseline workspace lock", manifest["before"].get("workspace_lock_sha256"), digest(gate_lock))
    expected_workspace = baseline_workspace(profile.BASELINE_COMMIT, gate_lock)
    equal_map("baseline committed workspace closure", manifest["before"].get("workspace_source_sha256"), expected_workspace)


def digest_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def accounting_fields(group: list[dict[str, Any]]) -> tuple[str, ...]:
    samples: list[dict[str, Any]] = []
    for record in group:
        raw = record.get("samples")
        if not isinstance(raw, list) or len(raw) != 1 or not isinstance(raw[0], dict):
            raise AuditError("measurement group has malformed raw sample")
        samples.append(raw[0])
    required = (
        "alloc_calls", "requested_bytes", "released_bytes", "live_before",
        "live_after", "peak_live_delta", "work", "memory_retained",
        "reference_reads", "output_bytes",
    )
    fields = tuple(field for field in required if all(field in sample for sample in samples))
    if not fields:
        raise AuditError("measurement group exposes no complete accounting fields")
    return fields


def accounting_sets(group: list[dict[str, Any]], fields: tuple[str, ...]) -> dict[str, list[int]]:
    result: dict[str, list[int]] = {}
    for field in fields:
        values: list[int] = []
        for record in group:
            values.append(int(record["samples"][0][field]))
        result[field] = sorted(values)
    return result


def bootstrap_interval(before: list[float], after: list[float], *, seed: int, resamples: int) -> list[float]:
    if not before or not after or resamples <= 0:
        raise AuditError("cannot calculate a bootstrap interval from empty timing samples")
    rng = random.Random(seed)
    shifts = sorted(
        100.0 * (
            median(rng.choices(after, k=len(after)))
            / median(rng.choices(before, k=len(before)))
            - 1.0
        )
        for _ in range(resamples)
    )
    lower = max(0, int(resamples * 0.025))
    upper = min(resamples - 1, int(resamples * 0.975) - 1)
    return [shifts[lower], shifts[upper]]


def compare_groups(
    baseline: dict[tuple[str, str], list[dict[str, Any]]],
    candidate: dict[tuple[str, str], list[dict[str, Any]]],
    *,
    seed: int,
    resamples: int,
    threshold: float,
) -> tuple[list[dict[str, Any]], tuple[str, ...]]:
    matched = sorted(set(baseline) & set(candidate))
    if not matched:
        raise AuditError("baseline and candidate have no matched groups")
    fields_by_group: dict[tuple[str, str], tuple[str, ...]] = {}
    for key in matched:
        before, after = baseline[key], candidate[key]
        equal(f"matched group {key} sample count", len(after), len(before))
        equal(
            f"matched group {key} sample indices",
            sorted(int(row["sample_index"]) for row in after),
            sorted(int(row["sample_index"]) for row in before),
        )
        fields = accounting_fields(before)
        candidate_fields = accounting_fields(after)
        equal(f"matched group {key} accounting schema", candidate_fields, fields)
        equal(f"matched group {key} accounting sets", accounting_sets(after, fields), accounting_sets(before, fields))
        before_checksums = sorted(int(row["samples"][0]["checksum"]) for row in before)
        after_checksums = sorted(int(row["samples"][0]["checksum"]) for row in after)
        equal(f"matched group {key} output checksums", after_checksums, before_checksums)
        fields_by_group[key] = fields

    comparisons: list[dict[str, Any]] = []
    for index, key in enumerate(matched):
        before, after = baseline[key], candidate[key]
        before_time = [float(row["elapsed_ns_p50"]) / int(row["repeat"]) for row in before]
        after_time = [float(row["elapsed_ns_p50"]) / int(row["repeat"]) for row in after]
        before_median = median(before_time)
        after_median = median(after_time)
        if before_median <= 0:
            raise AuditError(f"matched group {key} has nonpositive baseline timing")
        latency_delta = 100.0 * (after_median / before_median - 1.0)
        before_rss = median(int(row["rss_kib"]) for row in before)
        after_rss = median(int(row["rss_kib"]) for row in after)
        if before_rss <= 0:
            raise AuditError(f"matched group {key} has nonpositive baseline RSS")
        rss_delta = 100.0 * (after_rss / before_rss - 1.0)
        reasons: list[str] = []
        if latency_delta > threshold:
            reasons.append("latency")
        if rss_delta > threshold:
            reasons.append("rss")
        comparisons.append(
            {
                "case": key[0],
                "phase": key[1],
                "baseline_samples": len(before),
                "candidate_samples": len(after),
                "baseline_ns_per_repeat": before_median,
                "candidate_ns_per_repeat": after_median,
                "latency_delta_percent": latency_delta,
                "latency_bootstrap_95_percent": bootstrap_interval(
                    before_time,
                    after_time,
                    seed=seed + index,
                    resamples=resamples,
                ),
                "baseline_rss_kib": before_rss,
                "candidate_rss_kib": after_rss,
                "rss_delta_percent": rss_delta,
                "rss_delta_kib": after_rss - before_rss,
                "review_trigger": bool(reasons),
                "review_reasons": reasons,
                "accounting_fields": list(fields_by_group[key]),
            }
        )
    all_fields = sorted({field for fields in fields_by_group.values() for field in fields})
    return comparisons, tuple(all_fields)


def main() -> int:
    try:
        receipts.require_capture_inputs()
        if not RESULTS.is_dir():
            raise AuditError(f"missing retained performance results: {RESULTS}")
        profile_before = load(RESULTS / "profile-inputs-before.json")
        profile_after = load(RESULTS / "profile-inputs-after.json")
        if not isinstance(profile_before, dict) or not profile_before:
            raise AuditError("performance profile input receipt is empty or malformed")
        equal_map("profile input before/after", profile_after, profile_before)
        equal_map("current profile input", profile.profile_input_snapshot(), profile_before)
        summary = load(RESULTS / "capture-summary.json")
        equal_map("summary profile before", summary.get("profile_input_sha256_before"), profile_before)
        equal_map("summary profile after", summary.get("profile_input_sha256_after"), profile_after)
        cleanup = summary.get("cleanup")
        if not isinstance(cleanup, dict) or cleanup.get("baseline_removed") is not True:
            raise AuditError("baseline worktree cleanup receipt is not true")
        freeze = load(GATES / "freeze.json")
        gate_before = load(GATES / "source-before.json")
        if freeze.get("base_commit") != profile.BASELINE_COMMIT:
            raise AuditError("performance freeze base commit differs from profile baseline")
        gate_lock = GATES / "Cargo.lock"
        if not gate_lock.is_file() or digest(gate_lock) != profile.GATE_LOCK_SHA256:
            raise AuditError("retained gate Cargo.lock does not match the performance profile")
        captures = summary.get("captures")
        if not isinstance(captures, list) or len(captures) != 2:
            raise AuditError("retained capture summary must contain exactly one baseline/candidate pair")
        expected_labels = {f"baseline-{profile.BASELINE_COMMIT}", "candidate-final"}
        observed_labels = {capture.get("label") for capture in captures if isinstance(capture, dict)}
        equal("retained capture labels", observed_labels, expected_labels)
        warmups = summary.get("warmups_per_child")
        samples = summary.get("samples_per_group")
        if not isinstance(warmups, int) or warmups < 0 or not isinstance(samples, int) or samples <= 0:
            raise AuditError("retained capture warmup/sample settings are malformed")
        profile_contract = profile.CONTRACT_SHA256
        if not isinstance(profile_contract, str) or not profile_contract:
            raise AuditError("performance profile contract hash is not frozen")
        contract = HERE / "contract.md"
        if not contract.is_file() or digest(contract) != profile_contract:
            raise AuditError("performance profile contract hash is stale")
        if summary.get("contract_sha256") != profile_contract:
            raise AuditError("capture summary contract hash is stale")
        by_label = {capture["label"]: capture for capture in captures}
        baseline_capture = by_label[f"baseline-{profile.BASELINE_COMMIT}"]
        candidate_capture = by_label["candidate-final"]
        equal("capture summary controls", summary.get("controls"), list(profile.MATCHED_CONTROL_CASES))
        equal("capture summary phases", summary.get("phases"), list(profile.PHASES))
        candidate_preflight_gate = summary.get("candidate_preflight_gate")
        if not isinstance(candidate_preflight_gate, dict):
            raise AuditError("candidate preflight gate receipt is absent")
        equal("candidate preflight status", candidate_preflight_gate.get("status"), "ok")
        equal(
            "candidate preflight case count",
            candidate_preflight_gate.get("cases"),
            len(receipts.expected_candidate_cases()),
        )
        equal(
            "candidate preflight profile",
            candidate_preflight_gate.get("profile_input_sha256"),
            profile_before,
        )
        preflight_data = candidate_preflight_gate.get("preflight")
        if not isinstance(preflight_data, dict):
            raise AuditError("candidate preflight detail is malformed")
        preflight_relative = candidate_preflight_gate.get("output_dir")
        if not isinstance(preflight_relative, str):
            raise AuditError("candidate preflight output directory is absent")
        if preflight_relative.startswith("results/"):
            preflight_relative = preflight_relative[len("results/"):]
        preflight_dir = safe_relative(RESULTS, preflight_relative, "candidate preflight directory")
        if not preflight_dir.is_dir():
            raise AuditError(f"candidate preflight directory is absent: {preflight_dir}")
        receipts.verify_preflight(preflight_dir, list(receipts.expected_candidate_cases()))
        baseline_records, baseline_manifest, baseline_env = verify_capture(
            baseline_capture,
            profile_hashes=profile_before,
            warmups=warmups,
            samples=samples,
        )
        candidate_records, candidate_manifest, candidate_env = verify_capture(
            candidate_capture,
            profile_hashes=profile_before,
            warmups=warmups,
            samples=samples,
        )
        verify_baseline_source(baseline_manifest, gate_lock)
        excluded_gate_workspace_paths = verify_candidate_source(
            candidate_manifest, freeze, gate_before, gate_lock
        )
        verify_gate_payloads(gate_before, freeze, excluded_gate_workspace_paths)
        for field in ("rustc_verbose", "cargo", "libc", "rustflags", "contract_sha256"):
            equal(f"matched runtime {field}", baseline_env.get(field), candidate_env.get(field))
        equal("baseline environment contract", baseline_env.get("contract_sha256"), profile_contract)
        equal("candidate environment contract", candidate_env.get("contract_sha256"), profile_contract)
        equal(
            "matched harness identity",
            baseline_manifest["before"].get("harness_sha256"),
            candidate_manifest["before"].get("harness_sha256"),
        )
        equal(
            "matched profile identity",
            baseline_manifest["before"].get("profile_input_sha256"),
            candidate_manifest["before"].get("profile_input_sha256"),
        )
        baseline_groups = grouped(baseline_records)
        candidate_groups = grouped(candidate_records)
        matched_expected = {
            (case, phase)
            for case in profile.MATCHED_CONTROL_CASES
            for phase in profile.PHASES
        }
        equal("matched control groups", set(baseline_groups), matched_expected)
        if not matched_expected.issubset(candidate_groups):
            raise AuditError("candidate is missing one or more matched control groups")
        bootstrap = summary.get("bootstrap")
        if isinstance(bootstrap, dict):
            seed_value = bootstrap.get("seed")
            resamples_value = bootstrap.get("resamples")
            threshold_value = bootstrap.get("review_threshold_percent")
        else:
            seed_value = summary.get("bootstrap_seed")
            resamples_value = summary.get("bootstrap_resamples")
            threshold_value = summary.get("review_threshold_percent")
        seed = (
            int(seed_value)
            if isinstance(seed_value, int)
            else int.from_bytes(
                hashlib.sha256(
                    f"{profile.BASELINE_COMMIT}:{profile_contract}".encode("utf-8")
                ).digest()[:8],
                "big",
            )
        )
        resamples = int(resamples_value) if isinstance(resamples_value, int) else 5000
        threshold = float(threshold_value) if isinstance(threshold_value, (int, float)) else 5.0
        if resamples <= 0 or not (threshold >= 0):
            raise AuditError("bootstrap or review threshold settings are invalid")
        comparisons, accounting = compare_groups(
            baseline_groups,
            candidate_groups,
            seed=seed,
            resamples=resamples,
            threshold=threshold,
        )
        candidate_only = sorted(set(candidate_groups) - set(baseline_groups))
        result = {
            "status": "verified; individual performance flags retained",
            "verified_complete_samples": len(baseline_records) + len(candidate_records),
            "baseline_commit": profile.BASELINE_COMMIT,
            "candidate_label": candidate_capture["label"],
            "matched_control_groups": len(matched_expected),
            "candidate_only_groups": len(candidate_only),
            "candidate_only": [{"case": case, "phase": phase} for case, phase in candidate_only],
            "gate_workspace_payloads_excluded_from_profiler_closure": excluded_gate_workspace_paths,
            "accounting_fields": list(accounting),
            "accounting_sets_unchanged": True,
            "bootstrap": {
                "seed": seed,
                "resamples": resamples,
                "method": "independent unpaired median ratios; percentile interval",
                "review_threshold_percent": threshold,
            },
            "captures": [
                {
                    "label": "single-complete-pair",
                    "baseline": baseline_capture["label"],
                    "candidate": candidate_capture["label"],
                    "baseline_samples": len(baseline_records),
                    "candidate_samples": len(candidate_records),
                    "matched_groups": len(comparisons),
                    "manifest_sha256": {
                        baseline_capture["label"]: digest(RESULTS / baseline_capture["label"] / "source-manifest.json"),
                        candidate_capture["label"]: digest(RESULTS / candidate_capture["label"] / "source-manifest.json"),
                    },
                    "comparisons": comparisons,
                }
            ],
        }
        output = HERE / "root-performance-audit.json"
        output.write_text(json.dumps(result, indent=2, sort_keys=True) + "\n", encoding="utf-8")
        print(
            json.dumps(
                {
                    "status": "ok",
                    "samples": result["verified_complete_samples"],
                    "review_triggers": sum(1 for row in comparisons if row["review_trigger"]),
                },
                sort_keys=True,
            )
        )
        return 0
    except (OSError, KeyError, TypeError, ValueError, subprocess.CalledProcessError, RuntimeError) as error:
        raise SystemExit(f"retained performance audit failed: {error}")


if __name__ == "__main__":
    raise SystemExit(main())
