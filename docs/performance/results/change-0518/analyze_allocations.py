"""Validate and compare the 0518 DOCX publication allocation captures.

The allocator probe is intentionally separate from the native timing lanes.
This analyzer keeps every successful CSV row and tagged JSON sample in the
report, validates the build/capture custody chain, and derives only the
publication-region allocation counters and peaks.  It does not turn an
instrumented allocation or elapsed-time observation into a native speedup
claim.
"""

from __future__ import annotations

import argparse
import csv
import hashlib
import json
import re
import sys
from pathlib import Path
from typing import Any


HERE = Path(__file__).resolve().parent
REPO_ROOT = HERE.parents[3]
PROBE = HERE / "allocator-probe"
PLAN = HERE / "candidate-plan.json"
SOURCE_BINDING = PROBE / "source-binding.json"
OBSERVER_SOURCE = REPO_ROOT / "tools/perf-baseline/src/allocation_metrics.rs"
PROBE_CARGO_TOML = PROBE / "Cargo.toml"
PROBE_CARGO_LOCK = PROBE / "Cargo.lock"
WORKSPACE_CARGO_TOML = REPO_ROOT / "Cargo.toml"
WORKSPACE_CARGO_LOCK = REPO_ROOT / "Cargo.lock"
OBSERVER_REVISION = "serialized_region_peak_v3"
CLAIM_SCOPE = (
    "descriptive allocator evidence from an isolated standalone probe; the "
    "baseline/candidate comparison is valid within that identical probe, but "
    "absolute allocator counts and peaks do not represent normal native "
    "production-binary values or the workspace native release profile; no "
    "native timing or speedup claim"
)
ALLOCATION_SCOPE = (
    "publish_document_commit_to_stream method only; region begins immediately "
    "before the call and finishes immediately after return, before the caller "
    "drops the returned Snapshot"
)
TIMING_SCOPE = (
    "allocation counters are instrumented publication-region observations; "
    "JSON emission and allocator instrumentation are outside the native CSV "
    "timing claim"
)
EXPECTED_SOURCE_SHA256 = (
    "89ffdd564f6a76814fcc8b769f324557bd15a5201cb366ce9e75ade0a1c61181"
)
EXPECTED_GENERATED_LINES = 990
EXPECTED_PROBE_FILES = {
    "Cargo.lock",
    "Cargo.toml",
    "README.md",
    "generate.py",
    "source-binding.json",
    "source-diff.patch",
    "src/main.rs",
}
CASES = tuple(
    f"p{paragraphs}-k{replacements}-{source}-{mode}"
    for paragraphs in (128, 512)
    for replacements in (1, 8, 32)
    for source in ("owned", "file")
    for mode in ("repeated", "batch")
)
LANES = ("r1", "r2")
U64_MAX = (1 << 64) - 1
SHA256_RE = re.compile(r"^[0-9a-f]{64}$")
CASE_RE = re.compile(r"^p(128|512)-k(1|8|32)-(owned|file)-(repeated|batch)$")
BOOL_FIELDS = (
    "source_version_unchanged",
    "budget_managed",
    "memory_released",
    "objects_released",
    "input_monotonic",
    "work_monotonic",
    "semantic_ok",
    "raw_untouched_ok",
    "raw_untouched_member_payloads_ok",
    "output_exact_ok",
    "managed_preflight_forward_ok",
    "managed_preflight_inverse_ok",
    "forward_ok",
    "inverse_ok",
    "unmanaged_preflight_forward_ok",
    "unmanaged_preflight_inverse_ok",
)
PHASE_FIELDS = ("open_ns", "edit_ns", "commit_ns", "publish_ns", "drop_ns")
CSV_FIELDS = (
    "schema,version,api_path,comparison_scope,paragraphs,replacements,source,mode,"
    "fixture_sha256,fixture_bytes,fixture_name,expected_output_sha256,"
    "expected_output_bytes,repeat,ordinal,warmup,elapsed_ns,open_ns,edit_ns,"
    "commit_ns,publish_ns,drop_ns,output_bytes,output_sha256,"
    "source_version_before_id,source_version_before_revision,"
    "source_version_after_id,source_version_after_revision,source_version_unchanged,"
    "source_read_calls,source_requested_bytes,source_returned_bytes,"
    "source_zero_length_calls,cache_before_cold_loads,cache_before_successful_loads,"
    "cache_before_hits,cache_live_cold_loads,cache_live_successful_loads,"
    "cache_live_hits,cache_live_retained_bytes,cache_live_retained_entries,"
    "cache_live_in_flight_loads,budget_managed,budget_before_memory,"
    "budget_live_memory,budget_after_memory,budget_before_input,budget_live_input,"
    "budget_after_input,budget_before_output,budget_live_output,budget_after_output,"
    "budget_before_objects,budget_live_objects,budget_after_objects,"
    "budget_before_work,budget_live_work,budget_after_work,memory_released,"
    "objects_released,"
    "input_monotonic,work_monotonic,semantic_ok,raw_untouched_ok,"
    "raw_untouched_member_payloads_ok,output_exact_ok,"
    "managed_preflight_forward_ok,managed_preflight_inverse_ok,forward_ok,"
    "inverse_ok,unmanaged_preflight_forward_ok,unmanaged_preflight_inverse_ok"
).split(",")
# The tuple above is assembled from the frozen CSV header.  Keep a guard here
# so a duplicated typo in this analyzer cannot silently weaken the schema gate.
if len(CSV_FIELDS) != len(set(CSV_FIELDS)):
    raise RuntimeError("CSV_FIELDS contains duplicate names")

ALLOCATION_FIELDS = (
    "allocation_calls",
    "deallocation_calls",
    "reallocation_calls",
    "failed_allocation_calls",
    "allocated_bytes",
    "deallocated_bytes",
    "live_bytes_before",
    "live_bytes_after",
    "peak_live_bytes_before",
    "peak_live_bytes_after",
    "region_peak_live_bytes",
)
COMPARISON_FIELDS = (
    "allocation_calls",
    "deallocation_calls",
    "reallocation_calls",
    "allocated_bytes",
    "deallocated_bytes",
    "live_bytes_before",
    "live_bytes_after",
    "peak_live_bytes_before",
    "peak_live_bytes_after",
    "region_peak_live_bytes",
    "incremental_peak_bytes",
)


class AnalysisError(ValueError):
    """An evidence artifact is malformed or fails a required invariant."""


def sha(path: Path) -> str:
    digest = hashlib.sha256()
    try:
        with path.open("rb") as stream:
            for block in iter(lambda: stream.read(1024 * 1024), b""):
                digest.update(block)
    except OSError as error:
        raise AnalysisError(f"cannot hash {path}: {error}") from error
    return digest.hexdigest()


def read_text(path: Path) -> str:
    try:
        return path.read_text(encoding="utf-8", errors="strict")
    except (OSError, UnicodeError) as error:
        raise AnalysisError(f"cannot read {path}: {error}") from error


def regular(path: Path, context: str, nonempty: bool = True) -> None:
    if path.is_symlink() or not path.is_file():
        raise AnalysisError(f"{context} is missing or not a regular file: {path}")
    if nonempty and path.stat().st_size == 0:
        raise AnalysisError(f"{context} is empty: {path}")


def no_duplicate_pairs(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for key, value in pairs:
        if key in result:
            raise AnalysisError(f"duplicate JSON key {key!r}")
        result[key] = value
    return result


def reject_constant(value: str) -> None:
    raise AnalysisError(f"non-finite JSON number {value!r}")


def load_json(path: Path, context: str) -> Any:
    regular(path, context)
    try:
        return json.loads(
            read_text(path),
            object_pairs_hook=no_duplicate_pairs,
            parse_constant=reject_constant,
        )
    except AnalysisError:
        raise
    except json.JSONDecodeError as error:
        raise AnalysisError(f"{context} is invalid JSON: {error}") from error


def digest(value: Any, context: str) -> str:
    if not isinstance(value, str) or SHA256_RE.fullmatch(value) is None:
        raise AnalysisError(f"{context} is not a lowercase SHA-256 digest")
    return value


def integer(value: Any, context: str, maximum: int = U64_MAX) -> int:
    if isinstance(value, bool) or not isinstance(value, int):
        raise AnalysisError(f"{context} must be an integer")
    if value < 0 or value > maximum:
        raise AnalysisError(f"{context} is outside the unsigned 64-bit range")
    return value


def csv_integer(row: dict[str, str], field: str, context: str) -> int:
    value = row.get(field)
    if value is None or value == "":
        raise AnalysisError(f"{context}.{field} is missing")
    try:
        parsed = int(value, 10)
    except ValueError as error:
        raise AnalysisError(f"{context}.{field} is not an integer") from error
    if parsed < 0:
        raise AnalysisError(f"{context}.{field} is negative")
    return parsed


def require(condition: bool, message: str) -> None:
    if not condition:
        raise AnalysisError(message)


def case_identity(name: str) -> dict[str, Any]:
    match = CASE_RE.fullmatch(name)
    require(match is not None, f"invalid case name {name!r}")
    paragraphs, replacements, source, mode = match.groups()
    return {
        "paragraphs": int(paragraphs),
        "replacements": int(replacements),
        "source": source,
        "mode": mode,
        "api_path": (
            "replace_paragraph_text"
            if mode == "repeated"
            else "replace_body_paragraph_texts"
        ),
    }


def locked_package_version(text: str, package: str, context: str) -> str:
    pattern = re.compile(
        rf"(?ms)^\[\[package\]\]\s+name = {re.escape(json.dumps(package))}\s+"
        r'version = "([^"]+)"'
    )
    matches = pattern.findall(text)
    require(len(matches) == 1, f"{context} must contain one {package} package")
    return matches[0]


def allocator_build_scope(root: Path) -> dict[str, Any]:
    for path, context in (
        (PROBE_CARGO_TOML, "standalone allocator probe Cargo.toml"),
        (PROBE_CARGO_LOCK, "standalone allocator probe Cargo.lock"),
        (WORKSPACE_CARGO_TOML, "workspace Cargo.toml"),
        (WORKSPACE_CARGO_LOCK, "workspace Cargo.lock"),
    ):
        regular(path, context)
    probe_manifest = read_text(PROBE_CARGO_TOML)
    workspace_manifest = read_text(WORKSPACE_CARGO_TOML)
    require(
        re.search(r"(?m)^\[profile\.release\]\s*$", probe_manifest) is None,
        "standalone allocator probe unexpectedly defines [profile.release]",
    )
    require(
        re.search(r"(?m)^\[profile\.release\]\s*$", workspace_manifest) is not None,
        "workspace Cargo.toml omits [profile.release]",
    )
    require(
        re.search(r"(?m)^\s*lto\s*=\s*true\s*$", workspace_manifest) is not None,
        "workspace release profile does not enable LTO",
    )
    require(
        re.search(r'(?m)^\s*panic\s*=\s*"abort"\s*$', workspace_manifest) is not None,
        "workspace release profile does not set panic=abort",
    )
    probe_lock = read_text(PROBE_CARGO_LOCK)
    workspace_lock = read_text(WORKSPACE_CARGO_LOCK)
    probe_ryu = locked_package_version(probe_lock, "ryu", "standalone allocator probe Cargo.lock")
    workspace_ryu = locked_package_version(workspace_lock, "ryu", "workspace Cargo.lock")
    require(probe_ryu != workspace_ryu, "standalone and workspace ryu versions unexpectedly match")
    return {
        "kind": "standalone_allocator_probe_release_binary",
        "standalone_probe": {
            "manifest": relative(PROBE_CARGO_TOML, root),
            "manifest_sha256": sha(PROBE_CARGO_TOML),
            "lockfile": relative(PROBE_CARGO_LOCK, root),
            "lockfile_sha256": sha(PROBE_CARGO_LOCK),
            "dependency_resolution": "independent standalone lockfile",
            "dependency_examples": {"ryu": probe_ryu},
            "release_profile": {
                "profile_section": "no [profile.release]; Cargo defaults",
                "lto": "off (Cargo default)",
                "panic": "unwind (Cargo default)",
            },
        },
        "workspace_native": {
            "manifest": relative(WORKSPACE_CARGO_TOML, root),
            "manifest_sha256": sha(WORKSPACE_CARGO_TOML),
            "lockfile": relative(WORKSPACE_CARGO_LOCK, root),
            "lockfile_sha256": sha(WORKSPACE_CARGO_LOCK),
            "dependency_examples": {"ryu": workspace_ryu},
            "release_profile": {"lto": "true", "panic": "abort"},
        },
        "comparison_validity": (
            "Both allocator variants use the identical frozen probe source, "
            "standalone Cargo.toml, standalone Cargo.lock, and standalone "
            "release defaults; their baseline/candidate comparison is valid "
            "within this probe."
        ),
        "absolute_count_limit": (
            "Do not interpret absolute allocator counts or peaks as normal "
            "native production-binary counts or as measurements under the "
            "workspace native release profile."
        ),
    }


def relative(path: Path, root: Path) -> str:
    try:
        return path.relative_to(root).as_posix()
    except ValueError:
        return str(path)


def stage_build(stage: str, root: Path, warnings: list[dict[str, Any]]) -> dict[str, Any]:
    directory = root / f"allocator-{stage}"
    receipt_path = directory / "build-receipt.json"
    source_manifest_path = directory / "source-manifest.json"
    probe_manifest_path = directory / "probe-manifest.json"
    context = f"allocator-{stage}"
    result: dict[str, Any] = {
        "stage": stage,
        "directory": relative(directory, root),
        "receipt_path": relative(receipt_path, root),
        "source_manifest_path": relative(source_manifest_path, root),
        "probe_manifest_path": relative(probe_manifest_path, root),
        "valid": False,
        "issues": [],
    }
    try:
        receipt = load_json(receipt_path, f"{context} build receipt")
        source_manifest = load_json(source_manifest_path, f"{context} source manifest")
        probe_manifest = load_json(probe_manifest_path, f"{context} probe manifest")
        require(isinstance(receipt, dict), f"{context} build receipt is not an object")
        require(isinstance(source_manifest, dict), f"{context} source manifest is not an object")
        require(isinstance(probe_manifest, dict), f"{context} probe manifest is not an object")
        require(receipt.get("exit_code") == 0, f"{context} build exit_code is not zero")
        require(receipt.get("source_unchanged") is True, f"{context} source_unchanged is false")
        require(receipt.get("probe_unchanged") is True, f"{context} probe_unchanged is false")
        command = receipt.get("command")
        require(
            isinstance(command, list)
            and command[:3] == ["cargo", "build", "--release"]
            and "--locked" in command
            and "--features" in command
            and "allocator-metrics" in command,
            f"{context} build command is not the locked allocator probe build",
        )
        source_hash = digest(receipt.get("source_manifest_sha256"), f"{context}.source_manifest_sha256")
        probe_hash = digest(receipt.get("probe_manifest_sha256"), f"{context}.probe_manifest_sha256")
        require(
            sha(source_manifest_path) == source_hash,
            f"{context} source manifest receipt hash differs",
        )
        require(
            sha(probe_manifest_path) == probe_hash,
            f"{context} probe manifest receipt hash differs",
        )
        for name, value in source_manifest.items():
            digest(value, f"{context} source manifest {name}")
        for name, value in probe_manifest.items():
            digest(value, f"{context} probe manifest {name}")
        probe_prefix = "docs/performance/results/change-0518/allocator-probe/"
        actual_probe_names = {
            name[len(probe_prefix) :]
            for name in probe_manifest
            if name.startswith(probe_prefix)
        }
        require(
            actual_probe_names == EXPECTED_PROBE_FILES,
            f"{context} probe manifest file set differs",
        )
        for name, expected in probe_manifest.items():
            path = root.parent.parent.parent.parent / name
            regular(path, f"{context} probe manifest artifact {name}")
            require(sha(path) == expected, f"{context} probe artifact {name} hash differs")
        canonical_names = {
            "tools/perf-baseline/src/allocation_metrics.rs",
            "tools/perf-baseline/src/bin/support/counting_allocator.rs",
        }
        for name in canonical_names:
            require(name in source_manifest, f"{context} source manifest omits {name}")
            path = root.parent.parent.parent.parent / name
            regular(path, f"canonical observer source {name}")
            require(sha(path) == source_manifest[name], f"{context} {name} hash differs")
        binary_sha = digest(receipt.get("binary_sha256"), f"{context}.binary_sha256")
        binary_value = receipt.get("binary")
        require(isinstance(binary_value, str) and binary_value, f"{context}.binary is missing")
        binary_path = Path(binary_value)
        if binary_path.is_file() and not binary_path.is_symlink():
            require(sha(binary_path) == binary_sha, f"{context} binary hash differs")
        else:
            warnings.append(
                {
                    "kind": "binary_unavailable",
                    "stage": stage,
                    "path": binary_value,
                    "message": "build receipt binds the binary, but its temporary path is unavailable",
                }
            )
        result.update(
            {
                "valid": True,
                "receipt_sha256": sha(receipt_path),
                "source_manifest_sha256": source_hash,
                "probe_manifest_sha256": probe_hash,
                "binary": binary_value,
                "binary_sha256": binary_sha,
                "source_manifest_entries": len(source_manifest),
                "probe_manifest_entries": len(probe_manifest),
                "receipt": receipt,
                "canonical_source_sha256": {
                    name: source_manifest[name] for name in sorted(canonical_names)
                },
            }
        )
    except (AnalysisError, OSError) as error:
        issue = {"kind": "build_invalid", "stage": stage, "message": str(error)}
        result["issues"].append(issue)
        warnings.append(issue)
    return result


def validate_csv(path: Path, name: str) -> dict[str, Any]:
    identity = case_identity(name)
    regular(path, f"{name} CSV")
    context = f"{name} CSV"
    try:
        with path.open(newline="", encoding="utf-8") as stream:
            reader = csv.DictReader(stream)
            require(reader.fieldnames == CSV_FIELDS, f"{context} header differs from the frozen CSV schema")
            rows = list(reader)
    except (OSError, UnicodeError, csv.Error) as error:
        raise AnalysisError(f"{context} cannot be parsed: {error}") from error
    require(len(rows) == 1, f"{context} must contain exactly one data row")
    row = rows[0]
    require(all(value is not None for value in row.values()), f"{context} has a missing field")
    require(row["schema"] == "managed_paragraph_batch_perf_v1", f"{context}.schema differs")
    require(row["version"] == "1", f"{context}.version differs")
    for key in ("paragraphs", "replacements", "repeat", "ordinal"):
        csv_integer(row, key, context)
    require(
        (int(row["paragraphs"]), int(row["replacements"]), row["source"], row["mode"])
        == (identity["paragraphs"], identity["replacements"], identity["source"], identity["mode"]),
        f"{context} case identity differs",
    )
    require(row["api_path"] == identity["api_path"], f"{context}.api_path differs")
    require(
        row["comparison_scope"]
        == "descriptive_same_fixture_alternative_api_path_no_same_api_before_after_claim",
        f"{context}.comparison_scope differs",
    )
    require((int(row["repeat"]), int(row["ordinal"]), row["warmup"]) == (0, 0, "false"), f"{context} sample identity differs")
    for field in BOOL_FIELDS:
        require(row[field] == "true", f"{context}.{field} is not true")
    for field in PHASE_FIELDS + ("elapsed_ns",):
        csv_integer(row, field, context)
    require(
        int(row["elapsed_ns"]) >= sum(int(row[field]) for field in PHASE_FIELDS),
        f"{context} lifecycle elapsed is shorter than its phase sum",
    )
    require(row["expected_output_sha256"] == row["output_sha256"], f"{context} output hash oracle failed")
    require(
        int(row["expected_output_bytes"]) == int(row["output_bytes"]) > 0,
        f"{context} output byte oracle failed",
    )
    require(
        row["source_version_before_id"] == row["source_version_after_id"]
        and row["source_version_before_revision"] == row["source_version_after_revision"],
        f"{context} source version changed",
    )
    require(
        int(row["budget_before_memory"]) == int(row["budget_after_memory"]) == 0
        and int(row["budget_before_objects"]) == int(row["budget_after_objects"]) == 0,
        f"{context} memory/object budget release oracle failed",
    )
    require(
        int(row["budget_before_input"]) <= int(row["budget_live_input"]) <= int(row["budget_after_input"])
        and int(row["budget_before_work"]) <= int(row["budget_live_work"]) <= int(row["budget_after_work"]),
        f"{context} managed budget monotonicity oracle failed",
    )
    for field in (
        "fixture_sha256",
        "expected_output_sha256",
        "output_sha256",
        "source_version_before_id",
        "source_version_before_revision",
        "source_version_after_id",
        "source_version_after_revision",
    ):
        require(bool(row[field]), f"{context}.{field} is empty")
    return row


def validate_allocation(record: dict[str, Any], name: str) -> dict[str, Any]:
    context = f"{name} allocation sample"
    require(record.get("tag") == "allocationSample", f"{context} tag differs")
    require(record.get("scope") == "publish_document_commit_to_stream_method_only_before_returned_snapshot_drop", f"{context} scope differs")
    require(record.get("case") == name, f"{context} case differs")
    require(
        (record.get("repeat"), record.get("ordinal"), record.get("warmup")) == (0, 0, False),
        f"{context} identity differs",
    )
    sample = record.get("allocationSample")
    require(isinstance(sample, dict), f"{context}.allocationSample is missing")
    require(set(sample) == set(("status", "scope", *ALLOCATION_FIELDS)), f"{context} fields differ")
    require(sample.get("status") == "measured", f"{context} status is not measured")
    require(sample.get("scope") == "operation_global_system_allocator", f"{context} allocator scope differs")
    values = {field: integer(sample.get(field), f"{context}.{field}") for field in ALLOCATION_FIELDS}
    require(values["failed_allocation_calls"] == 0, f"{context} has failed allocation calls")
    require(
        values["live_bytes_before"] + values["allocated_bytes"] - values["deallocated_bytes"]
        == values["live_bytes_after"],
        f"{context} live-byte balance does not reconcile",
    )
    require(
        values["peak_live_bytes_before"] <= values["peak_live_bytes_after"],
        f"{context} absolute peak moved backwards",
    )
    require(
        values["region_peak_live_bytes"] >= max(values["live_bytes_before"], values["live_bytes_after"]),
        f"{context} region peak precedes a live-byte endpoint",
    )
    require(
        values["region_peak_live_bytes"] <= values["peak_live_bytes_after"],
        f"{context} region peak exceeds absolute peak",
    )
    values["incremental_peak_bytes"] = values["region_peak_live_bytes"] - values["live_bytes_before"]
    return values


def parse_stdout(path: Path, name: str) -> tuple[dict[str, Any], str]:
    regular(path, f"{name} stdout")
    raw = read_text(path)
    lines = [line for line in raw.splitlines() if line.strip()]
    require(len(lines) == 1, f"{name} stdout must contain exactly one tagged JSON line")
    try:
        value = json.loads(
            lines[0], object_pairs_hook=no_duplicate_pairs, parse_constant=reject_constant
        )
    except AnalysisError:
        raise
    except json.JSONDecodeError as error:
        raise AnalysisError(f"{name} stdout JSON is invalid: {error}") from error
    require(isinstance(value, dict), f"{name} stdout JSON is not an object")
    return value, raw


def validate_receipt(
    path: Path,
    name: str,
    stage: str,
    lane: str,
    build: dict[str, Any],
    root: Path,
) -> dict[str, Any]:
    context = f"{stage}/{lane}/{name} receipt"
    receipt = load_json(path, context)
    require(isinstance(receipt, dict), f"{context} is not an object")
    require(receipt.get("exit_code") == 0, f"{context} exit_code is not zero")
    require(receipt.get("source_unchanged") is True, f"{context} source_unchanged is false")
    require(receipt.get("cleanup_verified") is True, f"{context} cleanup_verified is false")
    require(receipt.get("binary_sha256") == build.get("binary_sha256"), f"{context} binary hash differs from build")
    require(receipt.get("source_manifest_sha256") == build.get("source_manifest_sha256"), f"{context} source manifest hash differs")
    require(receipt.get("probe_manifest_sha256") == build.get("probe_manifest_sha256"), f"{context} probe manifest hash differs")
    require(digest(receipt.get("candidate_plan_sha256"), f"{context}.candidate_plan_sha256") == sha(PLAN), f"{context} candidate plan hash differs")
    command = receipt.get("command")
    expected_prefix = ["/usr/bin/time", "-v", "taskset", "-c", "2"]
    require(isinstance(command, list) and command[:5] == expected_prefix, f"{context} command wrapper differs")
    require(len(command) == 24, f"{context} command length differs")
    identity = case_identity(name)
    expected = [
        *expected_prefix,
        build["binary"],
        "--paragraphs", str(identity["paragraphs"]),
        "--replacements", str(identity["replacements"]),
        "--source", identity["source"],
        "--mode", identity["mode"],
        "--samples", "1",
        "--warmups", "0",
        "--repeats", "1",
        "--artifact-dir", f"/tmp/litchi-goal-0518/corpora/alloc-{stage}-{lane}-{name}",
        "--output", str(root / f"alloc-{stage}-{lane}" / f"{name}.csv"),
    ]
    require(command == expected, f"{context} command arguments differ")
    artifacts = receipt.get("artifacts")
    require(isinstance(artifacts, dict), f"{context}.artifacts is missing")
    expected_names = {f"{name}.csv", f"{name}.stdout", f"{name}.stderr"}
    require(set(artifacts) == expected_names, f"{context} artifact set differs")
    for artifact_name, expected_hash in artifacts.items():
        artifact_path = path.parent / artifact_name
        regular(artifact_path, f"{context} artifact {artifact_name}")
        require(digest(expected_hash, f"{context}.artifacts.{artifact_name}") == sha(artifact_path), f"{context} artifact {artifact_name} hash differs")
    return receipt


def capture_lane(
    stage: str,
    lane: str,
    root: Path,
    build: dict[str, Any],
    warnings: list[dict[str, Any]],
) -> dict[str, Any]:
    directory = root / f"alloc-{stage}-{lane}"
    capture: dict[str, Any] = {
        "stage": stage,
        "lane": lane,
        "directory": relative(directory, root),
        "valid": False,
        "complete": False,
        "rows": [],
        "missing_cases": [],
        "issues": [],
    }
    for name in CASES:
        paths = {
            "csv": directory / f"{name}.csv",
            "stdout": directory / f"{name}.stdout",
            "stderr": directory / f"{name}.stderr",
            "receipt": directory / f"{name}.json",
        }
        entry: dict[str, Any] = {"case": name, "paths": {key: relative(value, root) for key, value in paths.items()}}
        try:
            receipt = validate_receipt(paths["receipt"], name, stage, lane, build, root)
            csv_row = validate_csv(paths["csv"], name)
            allocation_record, stdout_raw = parse_stdout(paths["stdout"], name)
            allocation_values = validate_allocation(allocation_record, name)
            entry.update(
                {
                    "valid": True,
                    "receipt": receipt,
                    "csv_row": csv_row,
                    "allocation_record": allocation_record,
                    "allocation_values": allocation_values,
                    "stdout_raw": stdout_raw,
                    "receipt_sha256": sha(paths["receipt"]),
                    "stderr_sha256": sha(paths["stderr"]),
                    "artifact_sha256": {
                        key: sha(value) for key, value in paths.items() if key != "receipt"
                    },
                }
            )
            capture["rows"].append(entry)
        except (AnalysisError, OSError) as error:
            issue = {
                "kind": "capture_invalid",
                "stage": stage,
                "lane": lane,
                "case": name,
                "message": str(error),
            }
            entry.update({"valid": False, "issue": issue})
            capture["rows"].append(entry)
            capture["issues"].append(issue)
            warnings.append(issue)
    capture["complete"] = len(capture["rows"]) == len(CASES) and all(row.get("valid") for row in capture["rows"])
    capture["valid"] = capture["complete"]
    return capture


def mean(values: list[int]) -> float:
    require(bool(values), "cannot average an empty metric vector")
    return sum(values) / len(values)


def percent_change(before: float, after: float) -> float | None:
    if before == 0:
        return None
    return (after / before - 1.0) * 100.0


def comparisons(captures: dict[str, Any]) -> list[dict[str, Any]]:
    by_stage_lane_case: dict[tuple[str, str, str], dict[str, Any]] = {}
    for stage in ("baseline", "candidate"):
        for lane in LANES:
            capture = captures.get(f"{stage}-{lane}")
            if not isinstance(capture, dict):
                continue
            for row in capture.get("rows", []):
                if row.get("valid"):
                    by_stage_lane_case[(stage, lane, row["case"])] = row
    result = []
    for name in CASES:
        item: dict[str, Any] = {"case": name, "baseline": {}, "candidate": {}, "comparison": {}}
        for stage in ("baseline", "candidate"):
            for lane in LANES:
                row = by_stage_lane_case.get((stage, lane, name))
                item[stage][lane] = (
                    {"allocation_values": row["allocation_values"], "path": row["paths"]}
                    if row is not None
                    else None
                )
        baseline_values = [
            by_stage_lane_case[("baseline", lane, name)]["allocation_values"]
            for lane in LANES
            if ("baseline", lane, name) in by_stage_lane_case
        ]
        candidate_values = [
            by_stage_lane_case[("candidate", lane, name)]["allocation_values"]
            for lane in LANES
            if ("candidate", lane, name) in by_stage_lane_case
        ]
        if len(baseline_values) == 2 and len(candidate_values) == 2:
            for field in COMPARISON_FIELDS:
                before = [values[field] for values in baseline_values]
                after = [values[field] for values in candidate_values]
                before_mean = mean(before)
                after_mean = mean(after)
                item["comparison"][field] = {
                    "baseline_r1": before[0],
                    "baseline_r2": before[1],
                    "candidate_r1": after[0],
                    "candidate_r2": after[1],
                    "baseline_mean": before_mean,
                    "candidate_mean": after_mean,
                    "candidate_minus_baseline_percent": percent_change(before_mean, after_mean),
                    "per_repeat_percent": [percent_change(before[index], after[index]) for index in range(2)],
                }
        result.append(item)
    return result


def observer_report(root: Path, builds: dict[str, Any]) -> dict[str, Any]:
    regular(OBSERVER_SOURCE, "canonical allocation observer source")
    source_text = read_text(OBSERVER_SOURCE)
    require(source_text.count(OBSERVER_REVISION) == 1, "canonical observer revision occurrence is not unique")
    source_hash = sha(OBSERVER_SOURCE)
    revision_hash = hashlib.sha256(OBSERVER_REVISION.encode()).hexdigest()
    canonical_manifest_hashes = {
        stage: builds[stage].get("canonical_source_sha256", {})
        for stage in ("baseline", "candidate")
        if builds.get(stage, {}).get("valid")
    }
    return {
        "revision": OBSERVER_REVISION,
        "revision_sha256": revision_hash,
        "source": relative(OBSERVER_SOURCE, root),
        "source_sha256": source_hash,
        "build_manifest_hashes": canonical_manifest_hashes,
    }


def analyze(root: Path) -> dict[str, Any]:
    warnings: list[dict[str, Any]] = []
    binding = load_json(SOURCE_BINDING, "probe source binding")
    require(isinstance(binding, dict), "probe source binding is not an object")
    require(binding.get("source_sha256") == EXPECTED_SOURCE_SHA256, "probe source binding source hash differs")
    require(binding.get("source_lines") == 907, "probe source binding source line count differs")
    require(binding.get("generated_lines") == EXPECTED_GENERATED_LINES, "probe generated line count differs")
    require(digest(binding.get("generated_sha256"), "probe generated_sha256") == sha(PROBE / "src/main.rs"), "probe generated source hash differs")
    plan_hash = digest(sha(PLAN), "candidate plan hash")
    builds = {stage: stage_build(stage, root, warnings) for stage in ("baseline", "candidate")}
    captures: dict[str, Any] = {}
    for stage in ("baseline", "candidate"):
        build = builds[stage]
        for lane in LANES:
            key = f"{stage}-{lane}"
            if not build.get("valid"):
                issue = {
                    "kind": "capture_skipped",
                    "stage": stage,
                    "lane": lane,
                    "message": "capture was not inspected because its allocator build evidence is invalid or absent",
                }
                warnings.append(issue)
                captures[key] = {
                    "stage": stage,
                    "lane": lane,
                    "directory": f"alloc-{stage}-{lane}",
                    "valid": False,
                    "complete": False,
                    "rows": [],
                    "missing_cases": list(CASES),
                    "issues": [issue],
                }
            else:
                captures[key] = capture_lane(stage, lane, root, build, warnings)
    complete = all(builds[stage].get("valid") for stage in ("baseline", "candidate")) and all(
        captures[f"{stage}-{lane}"].get("complete")
        for stage in ("baseline", "candidate")
        for lane in LANES
    )
    report = {
        "schema_version": 1,
        "status": "complete" if complete else "incomplete",
        "claim_scope": CLAIM_SCOPE,
        "allocation_scope": ALLOCATION_SCOPE,
        "timing_scope": TIMING_SCOPE,
        "allocator_build_scope": allocator_build_scope(root),
        "observer": observer_report(root, builds),
        "probe": {
            "source": binding.get("source"),
            "source_sha256": binding.get("source_sha256"),
            "source_lines": binding.get("source_lines"),
            "generated": binding.get("generated"),
            "generated_sha256": binding.get("generated_sha256"),
            "generated_lines": binding.get("generated_lines"),
            "probe_manifest_hashes": {
                stage: builds[stage].get("probe_manifest_sha256")
                for stage in ("baseline", "candidate")
                if builds[stage].get("probe_manifest_sha256")
            },
        },
        "candidate_plan_sha256": plan_hash,
        "builds": builds,
        "captures": captures,
        "comparisons": comparisons(captures),
        "warnings": warnings,
    }
    return report


def markdown(report: dict[str, Any]) -> str:
    lines = [
        "# DOCX publication allocation comparison",
        "",
        f"Status: **{report['status']}**.",
        "",
        report["claim_scope"],
        "",
        f"Allocation scope: {report['allocation_scope']}.",
        f"Timing scope: {report['timing_scope']}.",
        "",
        "The probe is bound to source "
        f"`{report['probe']['source_sha256']}` ({report['probe']['source_lines']} lines) "
        f"and generated source `{report['probe']['generated_sha256']}` "
        f"({report['probe']['generated_lines']} lines).  The canonical observer "
        f"revision is `{report['observer']['revision']}` with revision hash "
        f"`{report['observer']['revision_sha256']}` and source hash "
        f"`{report['observer']['source_sha256']}`.",
        "",
        "Build scope: the allocator binary is built from the standalone probe "
        f"manifest `{report['allocator_build_scope']['standalone_probe']['manifest']}` "
        f"and lockfile `{report['allocator_build_scope']['standalone_probe']['lockfile']}`. "
        "That lockfile resolves dependencies independently "
        f"(`ryu` {report['allocator_build_scope']['standalone_probe']['dependency_examples']['ryu']} "
        f"versus workspace {report['allocator_build_scope']['workspace_native']['dependency_examples']['ryu']}). "
        "The standalone manifest has no `[profile.release]`, so its release "
        "defaults use LTO off and panic unwinding; the workspace native release "
        "profile uses LTO=true and panic=abort.",
        "",
        report["allocator_build_scope"]["comparison_validity"],
        "",
        report["allocator_build_scope"]["absolute_count_limit"],
        "",
        "Each comparison uses two fresh one-row processes per case.  The "
        "absolute peak is `peak_live_bytes_after`; incremental peak is "
        "`region_peak_live_bytes - live_bytes_before`.  Full raw CSV rows and "
        "tagged JSON samples remain in the JSON report under `captures`.",
        "",
        "| Case | Allocation calls | Reallocation calls | Allocated bytes | Deallocated bytes | Region peak | Absolute peak | Incremental peak |",
        "| --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: |",
    ]
    for item in report["comparisons"]:
        values = item["comparison"]
        if not values:
            lines.append(f"| `{item['case']}` | unavailable | unavailable | unavailable | unavailable | unavailable | unavailable | unavailable |")
            continue

        def cell(field: str) -> str:
            value = values[field]
            before = value["baseline_mean"]
            after = value["candidate_mean"]
            change = value["candidate_minus_baseline_percent"]
            if change is None:
                delta = "n/a"
            else:
                delta = f"{change:+.2f}%"
            return f"{before:.1f} → {after:.1f} ({delta})"

        lines.append(
            f"| `{item['case']}` | {cell('allocation_calls')} | {cell('reallocation_calls')} | "
            f"{cell('allocated_bytes')} | {cell('deallocated_bytes')} | "
            f"{cell('region_peak_live_bytes')} | {cell('peak_live_bytes_after')} | "
            f"{cell('incremental_peak_bytes')} |"
        )
    lines.extend(["", "## Warnings", ""])
    if report["warnings"]:
        for warning in report["warnings"]:
            location = "/".join(
                str(warning[key])
                for key in ("stage", "lane", "case")
                if key in warning
            )
            prefix = f"`{location}`: " if location else ""
            lines.append(f"- {prefix}{warning['message']}")
    else:
        lines.append("No warnings.")
    lines.extend(["", "The allocation probe's instrumented elapsed time is excluded from native timing comparisons.", ""])
    return "\n".join(lines)


def write_output(path: Path, content: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(content, encoding="utf-8", newline="")


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--root", type=Path, default=HERE)
    parser.add_argument("--output", type=Path, default=HERE / "publication-allocation-comparison.json")
    parser.add_argument("--markdown", type=Path, default=HERE / "publication-allocation-comparison.md")
    parser.add_argument("--check-only", action="store_true")
    parser.add_argument("--allow-incomplete", action="store_true")
    args = parser.parse_args()
    try:
        report = analyze(args.root.resolve())
    except (AnalysisError, OSError) as error:
        print(f"allocation analysis failed: {error}", file=sys.stderr)
        return 1
    if not args.check_only:
        write_output(args.output, json.dumps(report, indent=2, sort_keys=True, allow_nan=False) + "\n")
        write_output(args.markdown, markdown(report))
    print(
        f"allocation analysis {report['status']}: "
        f"{len(report['warnings'])} warning(s)",
        flush=True,
    )
    return 0 if report["status"] == "complete" or args.allow_incomplete else 1


if __name__ == "__main__":
    raise SystemExit(main())
