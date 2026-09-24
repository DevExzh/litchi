#!/usr/bin/env python3
"""Audit one retained ODS lookup performance pair without recapturing it.

The audit is intentionally retained-only.  It reconstructs the baseline
workspace from Git, checks the candidate selected and gate closures, validates
raw measurements and preflight reads against the frozen case matrix, and
compares matched accounting sets.  It never creates a checkout, starts Cargo,
or embeds historical case/sample/outcome counts.
"""

from __future__ import annotations

from collections import defaultdict
import hashlib
import json
from pathlib import Path, PurePosixPath
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
CASE_MATRIX = PERFORMANCE / "case-matrix.json"

sys.path.insert(0, str(PERFORMANCE))


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


def digest_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def equal(label: str, observed: Any, expected: Any) -> None:
    if observed != expected:
        raise AuditError(f"{label}: expected {expected!r}, observed {observed!r}")


def equal_map(label: str, observed: dict[str, Any], expected: dict[str, Any]) -> None:
    if not isinstance(observed, dict) or not isinstance(expected, dict):
        raise AuditError(f"{label}: source maps are malformed")
    differences = [path for path in sorted(set(observed) | set(expected)) if observed.get(path) != expected.get(path)]
    if differences:
        raise AuditError(f"{label}: {len(differences)} mismatches; first {differences[:5]}")


def safe_relative(directory: Path, value: Any, label: str) -> Path:
    if not isinstance(value, str) or not value or Path(value).is_absolute():
        raise AuditError(f"{label} is not a relative path: {value!r}")
    path = (directory / value).resolve()
    try:
        path.relative_to(directory.resolve())
    except ValueError as error:
        raise AuditError(f"{label} escapes its directory: {value!r}") from error
    return path


def baseline_config() -> dict[str, Any]:
    value = load(HERE / "baseline.json")
    if not isinstance(value, dict):
        raise AuditError("baseline.json is not an object")
    return value


CONFIG = baseline_config()
BASELINE_COMMIT = CONFIG.get("commit")
GATE_LOCK_SHA256 = CONFIG.get("gate_lock_sha256")
if not isinstance(BASELINE_COMMIT, str) or not isinstance(GATE_LOCK_SHA256, str):
    raise AuditError("baseline commit or gate lock hash is missing")


def import_modules():
    try:
        import run_profile as profile  # type: ignore
        import verify as receipts  # type: ignore
    except ImportError as error:
        raise AuditError(f"lookup performance modules are unavailable: {error}") from error
    return profile, receipts


def profile_contract(profile: Any) -> str:
    value = getattr(profile, "CONTRACT_SHA256", None)
    if not isinstance(value, str) or not value or value.startswith("PENDING"):
        raise AuditError("performance profile contract hash is not frozen")
    contract = HERE / "contract.md"
    if not contract.is_file() or digest(contract) != value:
        raise AuditError("performance profile contract hash is stale")
    return value


def matrix_contract(profile: Any) -> dict[str, Any]:
    matrix = load(CASE_MATRIX)
    if not isinstance(matrix, dict):
        raise AuditError("performance case matrix is not an object")
    equal("case matrix baseline", matrix.get("baseline_commit"), BASELINE_COMMIT)
    phases = matrix.get("phases")
    if not isinstance(phases, list) or not phases or not all(isinstance(item, str) and item for item in phases):
        raise AuditError("case matrix phases are malformed")
    controls = matrix.get("controls", matrix.get("matched_controls"))
    candidates = matrix.get("candidate_cases", matrix.get("lookup_cases", matrix.get("cases")))
    if not isinstance(controls, list) or not controls or not all(isinstance(item, str) and item for item in controls):
        raise AuditError("case matrix controls are malformed")
    if not isinstance(candidates, list) or not candidates or not all(isinstance(item, str) and item for item in candidates):
        raise AuditError("case matrix candidate cases are malformed")
    if len(controls) != len(set(controls)) or len(candidates) != len(set(candidates)):
        raise AuditError("case matrix contains duplicate cases")
    expected_scope = CONFIG.get("scope")
    matrix_scope = matrix.get("scope", matrix.get("functions"))
    if not isinstance(matrix_scope, list) or sorted(matrix_scope) != sorted(expected_scope):
        raise AuditError("case matrix function scope differs from baseline")
    expected_reads = matrix.get("expected_reference_reads", matrix.get("reference_reads"))
    if not isinstance(expected_reads, dict):
        raise AuditError("case matrix read expectations are absent")
    direct: dict[str, int] = {}
    for case, value in expected_reads.items():
        if case in {"default", "candidate_cases", "controls", "exceptions", "by_case"}:
            continue
        if isinstance(value, bool) or not isinstance(value, int) or value < 0:
            raise AuditError(f"case matrix read expectation is malformed: {case}")
        direct[case] = value
    exceptions = expected_reads.get("exceptions", {})
    if exceptions is not None:
        if not isinstance(exceptions, dict):
            raise AuditError("case matrix read exceptions are malformed")
        for case, value in exceptions.items():
            if isinstance(value, bool) or not isinstance(value, int) or value < 0:
                raise AuditError(f"case matrix read exception is malformed: {case}")
            direct[case] = value
    by_case = expected_reads.get("by_case", {})
    if by_case is not None:
        if not isinstance(by_case, dict):
            raise AuditError("case matrix by_case reads are malformed")
        for case, value in by_case.items():
            if isinstance(value, bool) or not isinstance(value, int) or value < 0:
                raise AuditError(f"case matrix by_case read expectation is malformed: {case}")
            direct[case] = value
    defaults: dict[str, int | None] = {}
    for name in ("default", "candidate_cases", "controls"):
        value = expected_reads.get(name)
        if value is not None:
            if isinstance(value, bool) or not isinstance(value, int) or value < 0:
                raise AuditError(f"case matrix {name} read default is malformed")
            defaults[name] = value
    all_cases = set(controls) | set(candidates)
    missing = sorted(case for case in all_cases if case not in direct and not (case in candidates and "candidate_cases" in defaults) and not (case in controls and "controls" in defaults) and "default" not in defaults)
    if missing:
        raise AuditError(f"case matrix omits read expectations: {missing[:5]}")
    return {
        "scope": list(matrix_scope),
        "phases": phases,
        "controls": controls,
        "candidates": candidates,
        "reads": direct,
        "defaults": defaults,
        "sha256": digest(CASE_MATRIX),
    }


def expected_reads(contract: dict[str, Any], case: str) -> int:
    if case in contract["reads"]:
        return int(contract["reads"][case])
    defaults = contract["defaults"]
    if case in contract["candidates"] and "candidate_cases" in defaults:
        return int(defaults["candidate_cases"])
    if case in contract["controls"] and "controls" in defaults:
        return int(defaults["controls"])
    if "default" in defaults:
        return int(defaults["default"])
    raise AuditError(f"no frozen read expectation for {case}")


def git_blob_hashes(commit: str) -> dict[str, str]:
    listing = subprocess.check_output(["git", "ls-tree", "-r", commit, "Cargo.toml", "crates"], cwd=REPO, text=True)
    blobs: dict[str, str] = {}
    for line in listing.splitlines():
        header, relative = line.split("\t", 1)
        if relative == "Cargo.toml" or relative.endswith((".rs", "Cargo.toml", "build.rs")):
            fields = header.split()
            if len(fields) != 3 or fields[1] != "blob":
                raise AuditError(f"invalid Git tree entry: {line!r}")
            blobs[relative] = fields[2]
    result: dict[str, str] = {}
    process = subprocess.Popen(["git", "cat-file", "--batch"], cwd=REPO, stdin=subprocess.PIPE, stdout=subprocess.PIPE)
    assert process.stdin is not None and process.stdout is not None
    try:
        for oid in sorted(set(blobs.values())):
            process.stdin.write((oid + "\n").encode())
            process.stdin.flush()
            header = process.stdout.readline().split()
            if len(header) != 3 or header[1] != b"blob":
                raise AuditError(f"invalid Git batch response for {oid}")
            data = process.stdout.read(int(header[2]))
            if process.stdout.read(1) != b"\n":
                raise AuditError(f"invalid Git batch separator for {oid}")
            result[oid] = digest_bytes(data)
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
    try:
        data = subprocess.check_output(["git", "show", f"{commit}:{relative}"], cwd=REPO, stderr=subprocess.DEVNULL)
    except subprocess.CalledProcessError:
        return None
    return digest_bytes(data)


def baseline_profile_paths(profile: Any, commit: str) -> set[str]:
    listing = subprocess.check_output(["git", "ls-tree", "-r", "--name-only", commit], cwd=REPO, text=True).splitlines()
    paths = set(profile.SOURCE_FILES)
    for pattern in profile.SOURCE_FILE_GLOBS:
        # Match pathlib's recursive glob semantics without a live checkout;
        # **/ includes zero directories (for example reference/iri.rs).
        paths.update(path for path in listing if PurePosixPath(path).full_match(pattern))
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
        case, phase = record.get("case"), record.get("phase")
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


def expected_cases_for_capture(capture: dict[str, Any], matrix: dict[str, Any]) -> list[str]:
    scope = capture.get("case_scope")
    if scope in {"matched-controls", "controls"}:
        return list(matrix["controls"])
    if scope in {"all-named-cases", "all-cases", "candidate"}:
        ordered = list(matrix["controls"])
        ordered.extend(case for case in matrix["candidates"] if case not in ordered)
        return ordered
    raise AuditError(f"unknown capture case scope: {scope!r}")


def verify_preflight(directory: Path, expected: list[str], matrix: dict[str, Any]) -> None:
    preflight = load(directory / "preflight.json")
    equal(f"{directory} preflight status", preflight.get("status"), "ok")
    stdout = safe_relative(directory, preflight.get("stdout"), f"{directory} preflight stdout")
    if not stdout.is_file():
        raise AuditError(f"{directory}: preflight stdout missing")
    reads = preflight.get("reference_reads")
    if not isinstance(reads, dict):
        raise AuditError(f"{directory}: preflight reference reads missing")
    equal(f"{directory} preflight cases", sorted(reads), sorted(expected))
    for case in expected:
        equal(f"{directory} preflight reads {case}", int(reads[case]), expected_reads(matrix, case))


def verify_records(directory: Path, records: list[dict[str, Any]], expected: list[str], phases: list[str], warmups: int, samples: int, matrix: dict[str, Any]) -> None:
    equal(f"{directory} row count", len(records), len(expected) * len(phases) * samples)
    groups: dict[tuple[str, str], list[dict[str, Any]]] = defaultdict(list)
    for record in records:
        key = (record.get("case"), record.get("phase"))
        if key[0] not in expected or key[1] not in phases:
            raise AuditError(f"{directory}: unexpected group {key}")
        groups[key].append(record)
        equal(f"{directory} {key} supported", record.get("supported"), True)
        equal(f"{directory} {key} warmups", record.get("warmups"), warmups)
        if int(record.get("elapsed_ns_p50", 0)) <= 0 or int(record.get("rss_kib", 0)) <= 0:
            raise AuditError(f"{directory} {key}: invalid elapsed/RSS")
        repeat = int(record.get("repeat", 0))
        if repeat <= 0:
            raise AuditError(f"{directory} {key}: invalid repeat")
        raw_samples = record.get("samples")
        if not isinstance(raw_samples, list) or len(raw_samples) != 1:
            raise AuditError(f"{directory} {key}: child did not emit one raw sample")
        sample = raw_samples[0]
        if not isinstance(sample, dict):
            raise AuditError(f"{directory} {key}: raw sample is malformed")
        for field in ("elapsed_ns", "alloc_calls", "dealloc_calls", "requested_bytes", "released_bytes", "live_before", "live_after", "peak_live_delta", "work", "memory_retained", "reference_reads", "output_bytes"):
            if isinstance(sample.get(field), bool) or int(sample.get(field, -1)) < 0:
                raise AuditError(f"{directory} {key}: invalid {field}")
        equal(f"{directory} {key} allocator balance", sample["live_before"], sample["live_after"])
        if int(sample["released_bytes"]) > int(sample["requested_bytes"]):
            raise AuditError(f"{directory} {key}: released bytes exceed requested")
        total_reads = int(record.get("reference_reads_p50", -1))
        normalized = int(record.get("reference_reads_per_repeat", -1))
        expected_read_count = expected_reads(matrix, str(key[0]))
        if key[0] in {"lookup-cancel-vlookup", "lookup-cancel-indirect"}:
            # The frozen harness shares one sticky cancellation token across
            # four repeats: only the first repeat can reach the resolver.
            equal(f"{directory} {key} cancellation repeats", repeat, 4)
            equal(f"{directory} {key} cancellation preflight reads", expected_read_count, 1)
            expected_total_reads = 1
        else:
            expected_total_reads = expected_read_count * repeat
        equal(f"{directory} {key} total reference reads", total_reads, expected_total_reads)
        equal(f"{directory} {key} normalized reference reads", normalized, expected_total_reads // repeat)
        equal(f"{directory} {key} raw reference reads", int(sample["reference_reads"]), expected_total_reads)
        input_bytes = int(record.get("input_bytes", -1))
        output_bytes = int(record.get("output_bytes_p50", -1))
        bytes_per_repeat = int(record.get("bytes_per_repeat_p50", -1))
        if input_bytes < 0 or output_bytes < 0:
            raise AuditError(f"{directory} {key}: invalid byte metrics")
        equal(f"{directory} {key} normalized byte accounting", bytes_per_repeat, input_bytes + output_bytes // repeat)
    equal(f"{directory} groups", set(groups), {(case, phase) for case in expected for phase in phases})
    for key, group in groups.items():
        equal(f"{directory} {key} sample indices", sorted(int(row["sample_index"]) for row in group), list(range(1, samples + 1)))
        if len({row.get("binary_sha256") for row in group}) != 1:
            raise AuditError(f"{directory} {key}: binary changed across samples")


def verify_manifest(directory: Path, profile_hashes: dict[str, str]) -> dict[str, Any]:
    manifest = load(directory / "source-manifest.json")
    before, after = manifest.get("before"), manifest.get("after")
    if not isinstance(before, dict) or not isinstance(after, dict):
        raise AuditError(f"{directory}: source manifest is malformed")
    for key in ("source_sha256", "workspace_source_sha256", "profile_input_sha256"):
        equal(f"{directory} {key} stable", before.get(key), after.get(key))
    equal(f"{directory} source stable", manifest.get("source_sha256_unchanged"), True)
    equal(f"{directory} workspace stable", manifest.get("workspace_source_sha256_unchanged"), True)
    equal(f"{directory} profile stable", manifest.get("profile_input_sha256_unchanged"), True)
    equal(f"{directory} profile inputs", before.get("profile_input_sha256"), profile_hashes)
    equal(f"{directory} git stable", manifest.get("git_head_unchanged"), True)
    harness = before.get("harness_sha256")
    if not isinstance(harness, dict) or not harness.get("Cargo.lock"):
        raise AuditError(f"{directory}: harness identity is missing")
    cleanup = load(directory / "target-cleanup.json")
    equal(f"{directory} target cleanup", cleanup.get("removed"), True)
    return manifest


def verify_capture(capture: dict[str, Any], profile: Any, receipts: Any, matrix: dict[str, Any], profile_hashes: dict[str, str], warmups: int, samples: int) -> tuple[list[dict[str, Any]], dict[str, Any], dict[str, Any]]:
    label = capture.get("label")
    if not isinstance(label, str) or not label:
        raise AuditError("capture has no label")
    directory = RESULTS / label
    expected = expected_cases_for_capture(capture, matrix)
    equal(f"{label} capture cases", capture.get("cases"), expected)
    records = receipts.rows(directory)
    equal(f"{label} capture record count", capture.get("records"), len(records))
    manifest = verify_manifest(directory, profile_hashes)
    environment = load(directory / "environment.json")
    equal(f"{label} environment cases", environment.get("cases"), expected)
    equal(f"{label} environment phases", environment.get("phases"), matrix["phases"])
    equal(f"{label} source git head", manifest["before"].get("git_head"), BASELINE_COMMIT)
    equal(f"{label} stable source git head", manifest["after"].get("git_head"), BASELINE_COMMIT)
    equal(f"{label} capture source git head", capture.get("source_git_head"), BASELINE_COMMIT)
    equal(f"{label} capture binary", capture.get("binary_sha256"), manifest.get("binary_sha256"))
    verify_records(directory, records, expected, matrix["phases"], warmups, samples, matrix)
    verify_preflight(directory, expected, matrix)
    verify_raw_paths(directory, records)
    return records, manifest, environment


def compiled_workspace_map(workspace: dict[str, Any]) -> dict[str, Any]:
    return {path: value for path, value in workspace.items() if path in {"Cargo.toml", "Cargo.lock"} or (path.startswith("crates/") and path.endswith((".rs", "Cargo.toml", "build.rs")))}


def verify_source_pair(candidate_manifest: dict[str, Any], freeze: dict[str, Any], gate_before: dict[str, Any], gate_lock: Path) -> list[str]:
    equal("candidate freeze base commit", freeze.get("base_commit"), BASELINE_COMMIT)
    selected = freeze.get("selected_files")
    if not isinstance(selected, dict) or not selected:
        raise AuditError("performance freeze selected source map is malformed")
    source_map = candidate_manifest.get("before", {}).get("source_sha256")
    equal_map("candidate selected source closure", source_map, selected)
    for relative, expected in selected.items():
        path = gate_lock if relative == "Cargo.lock" else REPO / relative
        if not path.is_file():
            raise AuditError(f"candidate selected source is absent after cleanup: {relative}")
        equal(f"candidate selected source {relative}", digest(path), expected)
    gate_workspace = gate_before.get("workspace_source_sha256")
    candidate_workspace = candidate_manifest.get("before", {}).get("workspace_source_sha256")
    if not isinstance(gate_workspace, dict) or not isinstance(candidate_workspace, dict):
        raise AuditError("candidate/gate workspace closure is malformed")
    expected_compiled = compiled_workspace_map(gate_workspace)
    equal_map("candidate compiled workspace closure", candidate_workspace, expected_compiled)
    equal("candidate workspace lock", candidate_manifest["before"].get("workspace_lock_sha256"), digest(gate_lock))
    return sorted(set(gate_workspace) - set(expected_compiled))


def verify_gate_only_payloads(gate_before: dict[str, Any], freeze: dict[str, Any], excluded: list[str]) -> None:
    workspace = gate_before.get("workspace_source_sha256")
    selected = freeze.get("selected_files")
    if not isinstance(workspace, dict) or not isinstance(selected, dict):
        raise AuditError("gate workspace or selected source map is malformed")
    for relative in excluded:
        expected = selected.get(relative, committed_digest(BASELINE_COMMIT, relative))
        if expected is None:
            raise AuditError(f"gate-only payload is absent from baseline Git: {relative}")
        equal(f"gate-only payload {relative}", workspace.get(relative), expected)


def verify_baseline_source(manifest: dict[str, Any], profile: Any, gate_lock: Path) -> None:
    source_map = manifest.get("before", {}).get("source_sha256")
    if not isinstance(source_map, dict):
        raise AuditError("baseline selected source closure is malformed")
    equal("baseline selected source path set", set(source_map), baseline_profile_paths(profile, BASELINE_COMMIT))
    for relative, observed in source_map.items():
        expected = digest(gate_lock) if relative == "Cargo.lock" else committed_digest(BASELINE_COMMIT, relative)
        equal(f"baseline selected source {relative}", observed, expected)
    equal_map("baseline committed workspace closure", manifest["before"].get("workspace_source_sha256"), baseline_workspace(BASELINE_COMMIT, gate_lock))


def accounting_fields(group: list[dict[str, Any]]) -> tuple[str, ...]:
    required = ("alloc_calls", "requested_bytes", "released_bytes", "live_before", "live_after", "peak_live_delta", "work", "memory_retained", "reference_reads", "output_bytes")
    fields: list[str] = []
    for field in required:
        if all(isinstance(row.get("samples"), list) and len(row["samples"]) == 1 and field in row["samples"][0] for row in group):
            fields.append(field)
    if not fields:
        raise AuditError("measurement group exposes no complete accounting fields")
    return tuple(fields)


def accounting_sets(group: list[dict[str, Any]], fields: tuple[str, ...]) -> dict[str, list[int]]:
    return {field: sorted(int(row["samples"][0][field]) for row in group) for field in fields}


def bootstrap_interval(before: list[float], after: list[float], *, seed: int, resamples: int) -> list[float]:
    if not before or not after or resamples <= 0:
        raise AuditError("cannot calculate bootstrap interval from empty samples")
    rng = random.Random(seed)
    shifts = sorted(100.0 * (median(rng.choices(after, k=len(after))) / median(rng.choices(before, k=len(before))) - 1.0) for _ in range(resamples))
    lower = max(0, int(resamples * 0.025))
    upper = min(resamples - 1, int(resamples * 0.975) - 1)
    return [shifts[lower], shifts[upper]]


def compare_groups(baseline: dict[tuple[str, str], list[dict[str, Any]]], candidate: dict[tuple[str, str], list[dict[str, Any]]], *, seed: int, resamples: int, threshold: float) -> tuple[list[dict[str, Any]], tuple[str, ...]]:
    matched = sorted(set(baseline) & set(candidate))
    if not matched:
        raise AuditError("baseline and candidate have no matched groups")
    comparisons: list[dict[str, Any]] = []
    all_fields: set[str] = set()
    for index, key in enumerate(matched):
        before, after = baseline[key], candidate[key]
        equal(f"matched group {key} sample count", len(after), len(before))
        equal(f"matched group {key} sample indices", sorted(int(row["sample_index"]) for row in after), sorted(int(row["sample_index"]) for row in before))
        fields = accounting_fields(before)
        equal(f"matched group {key} accounting schema", accounting_fields(after), fields)
        equal(f"matched group {key} accounting sets", accounting_sets(after, fields), accounting_sets(before, fields))
        equal(f"matched group {key} output checksums", sorted(int(row["samples"][0]["checksum"]) for row in after), sorted(int(row["samples"][0]["checksum"]) for row in before))
        all_fields.update(fields)
        before_time = [float(row["elapsed_ns_p50"]) / int(row["repeat"]) for row in before]
        after_time = [float(row["elapsed_ns_p50"]) / int(row["repeat"]) for row in after]
        before_median, after_median = median(before_time), median(after_time)
        if before_median <= 0:
            raise AuditError(f"matched group {key} has nonpositive baseline timing")
        before_rss, after_rss = median(int(row["rss_kib"]) for row in before), median(int(row["rss_kib"]) for row in after)
        if before_rss <= 0:
            raise AuditError(f"matched group {key} has nonpositive baseline RSS")
        latency_delta = 100.0 * (after_median / before_median - 1.0)
        rss_delta = 100.0 * (after_rss / before_rss - 1.0)
        reasons = [reason for reason, value in (("latency", latency_delta), ("rss", rss_delta)) if value > threshold]
        comparisons.append({"case": key[0], "phase": key[1], "baseline_samples": len(before), "candidate_samples": len(after), "baseline_ns_per_repeat": before_median, "candidate_ns_per_repeat": after_median, "latency_delta_percent": latency_delta, "latency_bootstrap_95_percent": bootstrap_interval(before_time, after_time, seed=seed + index, resamples=resamples), "baseline_rss_kib": before_rss, "candidate_rss_kib": after_rss, "rss_delta_percent": rss_delta, "rss_delta_kib": after_rss - before_rss, "review_trigger": bool(reasons), "review_reasons": reasons, "accounting_fields": list(fields)})
    return comparisons, tuple(sorted(all_fields))


def verify_retained_rows_with_matrix(directory: Path, records: list[dict[str, Any]], expected: list[str], phases: list[str], warmups: int, samples: int, matrix: dict[str, Any], receipts: Any) -> dict[str, Any]:
    if not hasattr(receipts, "verify_rows") or not hasattr(receipts, "verify_preflight"):
        raise AuditError("retained performance verifier lacks row/preflight validation")
    equal("retained verifier phases", list(receipts.PHASES), phases)
    # Use the captured verifier unchanged. Independent exact read-count checks
    # above already cover every case, including sticky cancellation repeats.
    try:
        receipts.verify_rows(directory, records, expected, warmups, samples)
        receipts.verify_preflight(directory, expected)
    except RuntimeError as error:
        raise AuditError(f"captured retained verifier failed: {error}") from error
    return {"captured_verifier_failure": None, "corrected": False, "matrix_sha256": matrix["sha256"]}


def main() -> int:
    try:
        profile, receipts = import_modules()
        if getattr(profile, "BASELINE_COMMIT", BASELINE_COMMIT) != BASELINE_COMMIT:
            raise AuditError("performance profile baseline differs from evidence baseline")
        if getattr(profile, "GATE_LOCK_SHA256", GATE_LOCK_SHA256) != GATE_LOCK_SHA256:
            raise AuditError("performance profile gate lock differs from evidence baseline")
        contract_sha = profile_contract(profile)
        matrix = matrix_contract(profile)
        if not RESULTS.is_dir():
            raise AuditError(f"retained performance results are absent: {RESULTS}")
        profile_before = load(RESULTS / "profile-inputs-before.json")
        profile_after = load(RESULTS / "profile-inputs-after.json")
        if not isinstance(profile_before, dict) or not profile_before:
            raise AuditError("profile input receipt is malformed")
        equal_map("profile inputs before/after", profile_after, profile_before)
        equal_map("current profile input", profile.profile_input_snapshot(), profile_before)
        summary = load(RESULTS / "capture-summary.json")
        equal_map("summary profile before", summary.get("profile_input_sha256_before"), profile_before)
        equal_map("summary profile after", summary.get("profile_input_sha256_after"), profile_after)
        equal("summary contract hash", summary.get("contract_sha256"), contract_sha)
        cleanup = summary.get("cleanup")
        if not isinstance(cleanup, dict) or cleanup.get("baseline_removed") is not True:
            raise AuditError("baseline worktree cleanup receipt is not true")
        captures = summary.get("captures")
        if not isinstance(captures, list) or len(captures) != 2:
            raise AuditError("retained capture summary must contain one baseline/candidate pair")
        by_label = {capture.get("label"): capture for capture in captures if isinstance(capture, dict)}
        expected_labels = {f"baseline-{BASELINE_COMMIT}", "candidate-final"}
        equal("retained capture labels", set(by_label), expected_labels)
        warmups, samples = summary.get("warmups_per_child"), summary.get("samples_per_group")
        if not isinstance(warmups, int) or warmups < 0 or not isinstance(samples, int) or samples <= 0:
            raise AuditError("capture warmup/sample settings are malformed")
        baseline_records, baseline_manifest, baseline_env = verify_capture(by_label[f"baseline-{BASELINE_COMMIT}"], profile, receipts, matrix, profile_before, warmups, samples)
        candidate_records, candidate_manifest, candidate_env = verify_capture(by_label["candidate-final"], profile, receipts, matrix, profile_before, warmups, samples)
        gate_lock = GATES / "Cargo.lock"
        if not gate_lock.is_file() or digest(gate_lock) != GATE_LOCK_SHA256:
            raise AuditError("retained gate lock is absent or stale")
        verify_baseline_source(baseline_manifest, profile, gate_lock)
        gate_before = load(GATES / "source-before.json")
        freeze = load(GATES / "freeze.json")
        excluded = verify_source_pair(candidate_manifest, freeze, gate_before, gate_lock)
        verify_gate_only_payloads(gate_before, freeze, excluded)
        for field in ("rustc_verbose", "cargo", "libc", "rustflags", "contract_sha256"):
            equal(f"matched runtime {field}", baseline_env.get(field), candidate_env.get(field))
        equal("matched harness identity", baseline_manifest["before"].get("harness_sha256"), candidate_manifest["before"].get("harness_sha256"))
        equal("matched profile identity", baseline_manifest["before"].get("profile_input_sha256"), candidate_manifest["before"].get("profile_input_sha256"))
        retained_validation = verify_retained_rows_with_matrix(RESULTS / "candidate-final", candidate_records, expected_cases_for_capture(by_label["candidate-final"], matrix), matrix["phases"], warmups, samples, matrix, receipts)
        # These are descriptive audit settings, not capture inputs or a
        # pre-registered performance acceptance test. Keep every comparison
        # and flag; never use a later interval to erase an observed regression.
        bootstrap = {"seed": 20260920, "resamples": 10000, "review_threshold_percent": 5.0}
        seed = bootstrap.get("seed")
        resamples = bootstrap.get("resamples")
        threshold = bootstrap.get("review_threshold_percent")
        if isinstance(seed, bool) or not isinstance(seed, int) or isinstance(resamples, bool) or not isinstance(resamples, int) or resamples <= 0 or not isinstance(threshold, (int, float)) or threshold < 0:
            raise AuditError("capture summary bootstrap settings are malformed")
        comparisons, accounting = compare_groups(grouped(baseline_records), grouped(candidate_records), seed=seed, resamples=resamples, threshold=float(threshold))
        result = {"status": "verified; individual performance flags retained", "verified_complete_samples": len(baseline_records) + len(candidate_records), "baseline_commit": BASELINE_COMMIT, "candidate_label": "candidate-final", "matched_control_groups": len(set(grouped(baseline_records)) & set(grouped(candidate_records))), "candidate_only_groups": len(set(grouped(candidate_records)) - set(grouped(baseline_records))), "case_matrix_sha256": matrix["sha256"], "retained_verifier_read_bound_correction": retained_validation, "gate_workspace_payloads_excluded_from_profiler_closure": excluded, "accounting_fields": list(accounting), "accounting_sets_unchanged": True, "comparisons": comparisons, "review_triggers": [row for row in comparisons if row["review_trigger"]], "source_pair_verified": True, "harness_pair_verified": True}
        result["descriptive_bootstrap_settings"] = bootstrap
        (HERE / "root-performance-audit.json").write_text(json.dumps(result, indent=2, sort_keys=True) + "\n", encoding="utf-8")
        print(json.dumps({"status": "ok", "samples": result["verified_complete_samples"], "review_triggers": len(result["review_triggers"])}, sort_keys=True))
        return 0
    except (OSError, KeyError, TypeError, ValueError, subprocess.CalledProcessError, RuntimeError) as error:
        raise SystemExit(f"retained lookup performance audit failed: {error}") from error


if __name__ == "__main__":
    raise SystemExit(main())
