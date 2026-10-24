#!/usr/bin/env python3
"""Validate paired XLSX raw-identity reports after a guarded ABBA run.

The Rust probe owns the raw lifecycle and fidelity assertions.  This small
validator binds every report to its expected ABBA leg, source revision,
binary route, and build artifact, and checks that both the allocator and
native routes exercised the same four-corpus workload.  It deliberately does
not summarize elapsed values or make a performance claim.
"""

from __future__ import annotations

import hashlib
import json
from pathlib import Path
import sys
from typing import Any


SCHEMA = "litchi.xlsx.raw-identity-lifecycle.v2"
SUMMARY_SCHEMA = "litchi.xlsx.raw-native-qualification.v1"
BASELINE_REVISION = "0befb112d96d2409e6f40c9abfeb4813665905ad"
ABBA_ORDER = ("before", "after", "after", "before")
ROUTES = ("alloc", "native")
CASE_NAMES = (
    "small_single",
    "large_multipart_single",
    "small_multi_cache",
    "large_multipart_multi_cache",
)
OPERATION_NAMES = (
    "slicer_insert",
    "timeline_insert",
    "slicer_replace",
    "timeline_replace",
)
FIDELITY_FIELDS = (
    "untouched_worksheet_bytes",
    "replay_same_serialized_bytes",
    "inverse_restores_source_bytes",
    "stale_patch_refused",
    "no_op_preserves_source_bytes",
)


class ValidationError(ValueError):
    """A report or protocol invariant failed."""


def _expect(condition: bool, message: str) -> None:
    if not condition:
        raise ValidationError(message)


def _sha256(path: Path) -> tuple[str, int]:
    _expect(not path.is_symlink(), f"binary path may not be a symlink: {path}")
    _expect(path.is_file(), f"binary is missing: {path}")
    digest = hashlib.sha256()
    size = 0
    with path.open("rb") as stream:
        while block := stream.read(1024 * 1024):
            digest.update(block)
            size += len(block)
    return digest.hexdigest(), size


def _read_json(path: Path) -> dict[str, Any]:
    _expect(not path.is_symlink(), f"report may not be a symlink: {path}")
    _expect(path.is_file(), f"report is missing: {path}")
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeDecodeError, json.JSONDecodeError) as error:
        raise ValidationError(f"cannot read report {path}: {error}") from error
    _expect(isinstance(value, dict), f"report root is not an object: {path}")
    return value


def _validate_samples(samples: Any, *, expected_status: str, location: str, iterations: int) -> None:
    _expect(isinstance(samples, list) and len(samples) == iterations, f"{location} sample count mismatch")
    for index, sample in enumerate(samples):
        _expect(isinstance(sample, dict), f"{location}[{index}] is not an object")
        _expect(sample.get("status") == expected_status, f"{location}[{index}] allocation status mismatch")
        _expect(
            sample.get("scope") == "operation_global_system_allocator",
            f"{location}[{index}] allocation scope mismatch",
        )
        if expected_status == "measured":
            for field in (
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
            ):
                _expect(
                    isinstance(sample.get(field), int) and sample[field] >= 0,
                    f"{location}[{index}] has invalid {field}",
                )


def _validate_report(
    path: Path,
    *,
    route: str,
    variant: str,
    source_revision: str,
    target_root: Path,
) -> dict[str, Any]:
    report = _read_json(path)
    _expect(report.get("schema") == SCHEMA, f"{path}: report schema mismatch")
    _expect(report.get("api_variant") == variant, f"{path}: API variant mismatch")
    _expect(report.get("source_revision") == source_revision, f"{path}: source revision mismatch")
    _expect(report.get("corpus_generator") == "litchi-xlsx-raw-identity-package-v2", f"{path}: corpus mismatch")
    _expect(report.get("iterations") == 5 and report.get("warmups") == 2, f"{path}: probe repetition mismatch")
    _expect(
        "owned open" in str(report.get("lifecycle_scope", ""))
        and "native identity" in str(report.get("lifecycle_scope", "")),
        f"{path}: lifecycle scope omits raw identity coverage",
    )

    binary = report.get("binary")
    _expect(isinstance(binary, dict), f"{path}: binary identity is missing")
    reported_binary_path = Path(str(binary.get("path", "")))
    _expect(not reported_binary_path.is_symlink(), f"{path}: binary path may not be a symlink")
    binary_path = reported_binary_path.resolve(strict=False)
    expected_binary = (target_root / "release" / f"xlsx-raw-identity-{route}").resolve(strict=False)
    _expect(binary_path == expected_binary, f"{path}: binary path is not the declared {route} artifact")
    digest, size = _sha256(binary_path)
    _expect(binary.get("sha256") == digest, f"{path}: binary SHA-256 mismatch")
    _expect(binary.get("bytes") == size and size > 0, f"{path}: binary size mismatch")

    cases = report.get("cases")
    _expect(isinstance(cases, list), f"{path}: cases are missing")
    _expect([case.get("name") for case in cases] == list(CASE_NAMES), f"{path}: corpus order mismatch")
    for case_index, case in enumerate(cases):
        location = f"{path}:cases[{case_index}]"
        _expect(isinstance(case, dict), f"{location} is not an object")
        for field in ("extra_parts", "extra_part_bytes", "cache_count"):
            _expect(isinstance(case.get(field), int) and case[field] >= 0, f"{location}.{field} is invalid")
        for phase in ("open", "first_read"):
            phase_value = case.get(phase)
            _expect(isinstance(phase_value, dict), f"{location}.{phase} is missing")
            elapsed = phase_value.get("elapsed_ns")
            _expect(
                isinstance(elapsed, list)
                and len(elapsed) == 5
                and all(isinstance(value, int) and value >= 0 for value in elapsed),
                f"{location}.{phase}.elapsed_ns is invalid",
            )
            _validate_samples(
                phase_value.get("allocation_samples"),
                expected_status="measured" if route == "alloc" else "unavailable",
                location=f"{location}.{phase}.allocation_samples",
                iterations=5,
            )
        operations = case.get("operations")
        _expect(isinstance(operations, list), f"{location}.operations is missing")
        _expect([operation.get("operation") for operation in operations] == list(OPERATION_NAMES), f"{location}: operation order mismatch")
        for operation_index, operation in enumerate(operations):
            operation_location = f"{location}.operations[{operation_index}]"
            _expect(isinstance(operation, dict), f"{operation_location} is not an object")
            for field in ("base_bytes", "output_bytes"):
                _expect(isinstance(operation.get(field), int) and operation[field] > 0, f"{operation_location}.{field} is invalid")
            for field in ("base_sha256", "output_sha256"):
                value = operation.get(field)
                _expect(isinstance(value, str) and len(value) == 64, f"{operation_location}.{field} is invalid")
            _expect(isinstance(operation.get("native_identity"), list) and operation["native_identity"], f"{operation_location}.native_identity is missing")
            _expect(
                all(isinstance(identity, str) and identity for identity in operation["native_identity"]),
                f"{operation_location}.native_identity is invalid",
            )
            _expect(
                all(operation.get(field) is True for field in FIDELITY_FIELDS),
                f"{operation_location}: fidelity guard failed",
            )
            _expect(
                isinstance(operation.get("elapsed_ns"), list)
                and len(operation["elapsed_ns"]) == 5
                and all(isinstance(value, int) and value >= 0 for value in operation["elapsed_ns"]),
                f"{operation_location}.elapsed_ns is invalid",
            )
            _validate_samples(
                operation.get("allocation_samples"),
                expected_status="measured" if route == "alloc" else "unavailable",
                location=f"{operation_location}.allocation_samples",
                iterations=5,
            )
    return {
        "path": str(path),
        "route": route,
        "variant": variant,
        "source_revision": source_revision,
        "sha256": _sha256(path)[0],
        "bytes": path.stat().st_size,
        "cases": len(cases),
        "fidelity_operations": len(cases) * len(OPERATION_NAMES),
    }


def validate_directory(
    output_dir: Path,
    baseline_revision: str,
    candidate_revision: str,
    before_target: Path,
    after_target: Path,
) -> dict[str, Any]:
    """Validate all eight reports produced by the paired four-leg ABBA run."""

    _expect(baseline_revision == BASELINE_REVISION, "the XLSX baseline must remain the original 0bef revision")
    _expect(
        len(candidate_revision) == 40
        and all(char in "0123456789abcdefABCDEF" for char in candidate_revision)
        and candidate_revision != BASELINE_REVISION,
        "candidate revision must be a distinct 40-character hexadecimal revision",
    )
    reports: list[dict[str, Any]] = []
    expected_paths: set[Path] = set()
    for index, leg in enumerate(ABBA_ORDER, start=1):
        variant = "original" if leg == "before" else "candidate"
        revision = baseline_revision if leg == "before" else candidate_revision
        target = before_target if leg == "before" else after_target
        for route in ROUTES:
            report_path = output_dir / f"xlsx-raw-{route}-{leg}-{index}.json"
            expected_paths.add(report_path)
            reports.append(
                _validate_report(
                    report_path,
                    route=route,
                    variant=variant,
                    source_revision=revision,
                    target_root=target,
                )
            )
    actual_paths = set(output_dir.glob("xlsx-raw-*.json"))
    _expect(actual_paths == expected_paths, "XLSX raw output set differs from the fixed ABBA report set")
    return {
        "schema": SUMMARY_SCHEMA,
        "abba_order": list(ABBA_ORDER),
        "source_revisions": {"before": baseline_revision, "after": candidate_revision},
        "routes": list(ROUTES),
        "reports": reports,
        "fidelity": {"reports": len(reports), "operations_per_report": len(CASE_NAMES) * len(OPERATION_NAMES)},
        "allocation_routes": {"alloc": "measured", "native": "unavailable"},
        "performance_claim": False,
    }


def main(argv: list[str]) -> int:
    if len(argv) != 5:
        print(
            "usage: xlsx_raw_identity_validator.py OUTPUT_DIR BASELINE_REVISION "
            "CANDIDATE_REVISION BEFORE_TARGET AFTER_TARGET",
            file=sys.stderr,
        )
        return 2
    output_dir = Path(argv[0])
    baseline_revision = argv[1]
    candidate_revision = argv[2]
    before_target = Path(argv[3])
    after_target = Path(argv[4])
    try:
        summary = validate_directory(
            output_dir,
            baseline_revision,
            candidate_revision,
            before_target,
            after_target,
        )
        summary_path = output_dir / "summary.json"
        _expect(not summary_path.exists() and not summary_path.is_symlink(), "summary.json already exists")
        with summary_path.open("x", encoding="utf-8") as stream:
            json.dump(summary, stream, indent=2, sort_keys=True)
            stream.write("\n")
        print(f"validated {len(summary['reports'])} XLSX raw reports")
    except (OSError, ValidationError) as error:
        print(f"XLSX raw qualification validation failed: {error}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
