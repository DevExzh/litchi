#!/usr/bin/env python3
"""Qualify the 0467 fixed-checkout ABBA experiment.

The original 0467 500-sample helper targets the superseded role-specific
worktrees.  This helper is additive and handles the fixed-checkout experiment:
both binaries are built and captured from the same checkout path, with the
role revision checked out at each capture boundary.  Four 500-sample rows are
retained: the XLSX primary row and three DOC rows kept as deliberate guards.

The helper validates the captured receipts and bindings, delegates the raw
report and ABBA statistic checks to ``perf_abba_summary``, records process-wide
GNU ``time`` RSS, and emits a proposal-only primary claim.  It never edits the
claim registry or treats the DOC guard rows as claim evidence.
"""

from __future__ import annotations

import argparse
import copy
import json
from pathlib import Path
import sys
from typing import Any, Iterable, Mapping, Sequence


ROOT = Path(__file__).resolve().parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))
REPO_ROOT = next(
    (candidate for candidate in ROOT.parents if (candidate / "tools").is_dir()),
    ROOT,
)
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

import analyze as normal  # noqa: E402
from tools import perf_abba_summary, perf_compare  # noqa: E402


SCHEMA = "litchi-0467-fixed-checkout-qualification-v1"
PROTOCOL_SCHEMAS = {
    "litchi-0467-fixed-path-protocol-v1",
    "litchi-0467-fixed-checkout-qualification-v1",
    "litchi-0467-fixed-qualification-protocol-v1",
}
PROTOCOL_CANDIDATES = (
    "protocol-fixed.json",
    "protocol-fixed500.json",
    "protocol-fixed-checkout.json",
)

FIXED_LANES = ("A1-fixed", "B1-fixed", "B2-fixed", "A2-fixed")
FIXED_LEGS = ("a1", "b1", "b2", "a2")
FIXED_ROLES = {
    "A1-fixed": "control",
    "A2-fixed": "control",
    "B1-fixed": "candidate",
    "B2-fixed": "candidate",
}
FIXED_BINDINGS = {
    "control": "control-fixed-binding.json",
    "candidate": "candidate-fixed-binding.json",
}

# The full fixed-checkout guard is accepted with either spelling so the
# capture driver may make the clean suffix explicit without changing the
# normal four-lane protocol.
FULL_FIXED_LANE_PAIRS = (
    ("A-full-fixed", "B-full-fixed"),
    ("A-full-fixed-clean", "B-full-fixed-clean"),
)

DOC_CASE = "doc_fresh_write_to"
XLSX_CASE = "xlsx_one_percent_commit_save"
FIXED_CASES = (DOC_CASE, XLSX_CASE)
DOC_SHAPES = ("tiny", "large", "payload-heavy")
XLSX_SHAPE = "dense-wide"
SAMPLES = 500
WARMUPS = 5
CPU = "2"
WORKERS = 1
REVIEW_THRESHOLD_PERCENT = 5.0
DRIFT_CEILINGS = {"p50": 5.0, "mean": 5.0, "p95": 10.0, "p99": 15.0}
STATISTICS = ("p50", "mean", "p95", "p99")
FULL_SAMPLES = 15
FULL_WARMUPS = 3
FULL_ROWS = 201
PACKAGE_CHANGE_ID = "0467-xlsx-cell-attributes-abba"
CLAIM_ID = "claim-0467-xlsx-cell-attributes"
EVIDENCE_ID = "abba-0467-xlsx-cell-attributes-fixed-checkout"
DOCUMENTATION = "docs/performance/changes/0467-xlsx-cell-attributes.md"

EXPECTED_ROW_SHAPES = {
    (DOC_CASE, shape) for shape in DOC_SHAPES
} | {(XLSX_CASE, XLSX_SHAPE)}


class QualificationError(normal.AnalysisError):
    """Raised when fixed-checkout evidence does not meet the protocol."""


def _primary_report_projection(
    report: Mapping[str, Any], label: str = "report"
) -> dict[str, Any]:
    """Return a claim-scoped report containing only the primary XLSX row.

    Fixed normal captures intentionally retain the three DOC guard rows.  The
    generic ABBA package verifier requires every summary row to be inside the
    registry claim scope, so packaging uses this derived one-row report for
    each ABBA leg.  The original four-row report remains the authenticated
    custody record in the qualification output.  Parallel-metrics entries are
    projected by the same result index so the standard report validator keeps
    its cardinality and case/corpus identity checks.
    """

    root = copy.deepcopy(_object(report, label))
    results = root.get("results")
    if not isinstance(results, list):
        raise QualificationError(f"{label}.results must be a list")
    selected_indices: list[int] = []
    for index, raw_row in enumerate(results):
        row = _object(raw_row, f"{label}.results[{index}]")
        corpus = _object(row.get("corpus"), f"{label}.results[{index}].corpus")
        if row.get("case") == XLSX_CASE and corpus.get("shape") == XLSX_SHAPE:
            selected_indices.append(index)
    if len(selected_indices) != 1:
        raise QualificationError(
            f"{label} must contain exactly one {XLSX_CASE}/{XLSX_SHAPE} row"
        )

    configuration = _object(root.get("configuration"), f"{label}.configuration")
    configuration["cases"] = [XLSX_CASE]
    root["configuration"] = configuration
    root["results"] = [results[index] for index in selected_indices]

    parallel_metrics = root.get("parallel_metrics")
    if parallel_metrics is not None:
        parallel_metrics = _object(
            parallel_metrics, f"{label}.parallel_metrics"
        )
        parallel_cases = parallel_metrics.get("cases")
        if not isinstance(parallel_cases, list) or len(parallel_cases) != len(results):
            raise QualificationError(
                f"{label}.parallel_metrics.cases must align with results"
            )
        parallel_metrics["cases"] = [
            parallel_cases[index] for index in selected_indices
        ]
        root["parallel_metrics"] = parallel_metrics
    return root


def _object(value: Any, location: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        raise QualificationError(f"{location} must be an object")
    return value


def _string(value: Any, location: str) -> str:
    try:
        return normal._string(value, location)
    except normal.AnalysisError as error:
        raise QualificationError(str(error)) from error


def _sha(value: Any, location: str) -> str:
    try:
        return normal._sha(value, location)
    except normal.AnalysisError as error:
        raise QualificationError(str(error)) from error


def _revision(value: Any, location: str) -> str:
    try:
        return normal._revision(value, location)
    except normal.AnalysisError as error:
        raise QualificationError(str(error)) from error


def _exact_list(value: Any, expected: Sequence[Any], location: str) -> None:
    if value != list(expected):
        raise QualificationError(f"{location} does not match {list(expected)!r}: {value!r}")


def _protocol_path(root: Path, requested: Path | None) -> Path:
    if requested is not None:
        path = requested.resolve()
        if not path.is_file():
            raise QualificationError(f"missing fixed-checkout protocol {path}")
        return path
    for name in PROTOCOL_CANDIDATES:
        path = root / name
        if path.is_file():
            return path
    raise QualificationError(
        "missing fixed-checkout protocol; pass --protocol"
    )


def _validate_protocol(path: Path) -> dict[str, Any]:
    protocol = _object(normal.load_json(path), str(path))
    if protocol.get("schema") not in PROTOCOL_SCHEMAS:
        raise QualificationError(f"{path}.schema is not a fixed-checkout protocol")
    if protocol.get("status") != "frozen":
        raise QualificationError(f"{path}.status must be frozen")
    _exact_list(protocol.get("roles"), ("control", "candidate"), f"{path}.roles")
    _exact_list(protocol.get("order"), FIXED_LANES, f"{path}.order")
    if protocol.get("samples") != SAMPLES or protocol.get("warmups") != WARMUPS:
        raise QualificationError(f"{path} must specify 500 samples and 5 warmups")
    if protocol.get("cpu") != 2 or protocol.get("workers") != WORKERS:
        raise QualificationError(f"{path} must specify CPU 2 and one worker")
    _exact_list(protocol.get("cases"), FIXED_CASES, f"{path}.cases")
    _exact_list(protocol.get("shapes"), (XLSX_SHAPE,), f"{path}.shapes")
    if protocol.get("expected_result_count") != len(EXPECTED_ROW_SHAPES):
        raise QualificationError(f"{path}.expected_result_count must be {len(EXPECTED_ROW_SHAPES)}")
    if protocol.get("shared_tree") != "/tmp/litchi-goal-0467/shared-tree":
        raise QualificationError(f"{path}.shared_tree must identify the fixed shared checkout")
    if protocol.get("flags") != "-C force-frame-pointers=yes -C force-unwind-tables=yes":
        raise QualificationError(f"{path}.flags do not match the fixed build")
    if protocol.get("debug_level") != 1 or protocol.get("toolchain") != "1.98.1":
        raise QualificationError(f"{path} toolchain/debug configuration is not fixed")
    _exact_list(
        protocol.get("initial_guards"), ("A-full-fixed", "B-full-fixed"),
        f"{path}.initial_guards",
    )
    if protocol.get("initial_guard_samples") != FULL_SAMPLES:
        raise QualificationError(f"{path}.initial_guard_samples must be {FULL_SAMPLES}")
    if protocol.get("initial_guard_warmups") != FULL_WARMUPS:
        raise QualificationError(f"{path}.initial_guard_warmups must be {FULL_WARMUPS}")
    if protocol.get("initial_guard_expected_rows") != FULL_ROWS:
        raise QualificationError(f"{path}.initial_guard_expected_rows must be {FULL_ROWS}")
    binding_files = protocol.get("role_binding_files")
    if binding_files != FIXED_BINDINGS:
        raise QualificationError(f"{path}.role_binding_files do not match fixed bindings")
    if protocol.get("review_trigger_percent") != REVIEW_THRESHOLD_PERCENT:
        raise QualificationError(f"{path}.review_trigger_percent must be 5")
    ceilings = protocol.get("strict_policy_drift_ceilings_percent")
    if ceilings != {key: int(value) for key, value in DRIFT_CEILINGS.items()}:
        raise QualificationError(f"{path}.strict_policy_drift_ceilings_percent is not 5/5/10/15")
    triggers = protocol.get("review_triggers")
    if triggers is not None:
        triggers = _object(triggers, f"{path}.review_triggers")
        if triggers.get("latency_percent") != REVIEW_THRESHOLD_PERCENT:
            raise QualificationError(f"{path}.review_triggers.latency_percent must be 5")
        if triggers.get("process_rss_percent") != REVIEW_THRESHOLD_PERCENT:
            raise QualificationError(f"{path}.review_triggers.process_rss_percent must be 5")
    full_lanes = protocol.get("full_guard_order")
    if full_lanes is not None and full_lanes not in [list(pair) for pair in FULL_FIXED_LANE_PAIRS]:
        raise QualificationError(f"{path}.full_guard_order is not a recognized fixed guard pair")
    return protocol


def _fixed_binding(root: Path, role: str) -> tuple[Path, dict[str, Any]]:
    if role not in FIXED_BINDINGS:
        raise QualificationError(f"unsupported fixed binding role {role!r}")
    path = root / FIXED_BINDINGS[role]
    if not path.is_file():
        raise QualificationError(f"missing fixed binding {path}")
    binding = _object(normal.load_json(path), str(path))
    if binding.get("role") != role:
        raise QualificationError(f"{path}.role does not match {role}")
    revision = _revision(binding.get("revision"), f"{path}.revision")
    binary = _sha(binding.get("binary_sha256"), f"{path}.binary_sha256")
    binary_bytes = binding.get("bytes")
    if isinstance(binary_bytes, bool) or not isinstance(binary_bytes, int) or binary_bytes <= 0:
        raise QualificationError(f"{path}.bytes must be a positive integer")
    receipt_name = _string(binding.get("build_receipt"), f"{path}.build_receipt")
    receipt_relative = Path(receipt_name)
    if receipt_relative.is_absolute() or ".." in receipt_relative.parts:
        raise QualificationError(f"{path}.build_receipt must be bundle-relative")
    receipt_path = path.parent / receipt_relative
    build_receipt = _object(normal.load_json(receipt_path), str(receipt_path))
    if build_receipt.get("schema") != "litchi-0467-fixed-build-v1":
        raise QualificationError(f"{receipt_path}.schema is not the fixed build receipt")
    receipt_digest = normal.raw_sha256(receipt_path)
    if receipt_digest != _sha(
        binding.get("build_receipt_sha256"), f"{path}.build_receipt_sha256"
    ):
        raise QualificationError(f"{path}.build_receipt_sha256 does not match its receipt")
    receipt_revision = _revision(build_receipt.get("revision"), f"{receipt_path}.revision")
    if receipt_revision != revision:
        raise QualificationError(f"{receipt_path}.revision does not match its binding")
    receipt_role = build_receipt.get("role")
    if receipt_role not in (role, f"{role}-fixed"):
        raise QualificationError(f"{receipt_path}.role does not match {role}")
    original_binding = _string(
        build_receipt.get("original_binding"), f"{receipt_path}.original_binding"
    )
    if original_binding != f"{role}-binding.json":
        raise QualificationError(f"{receipt_path}.original_binding does not match {role}")
    original_binding_path = path.parent / original_binding
    if not original_binding_path.is_file() or build_receipt.get("original_binding_sha256") != normal.raw_sha256(original_binding_path):
        raise QualificationError(f"{receipt_path}.original_binding_sha256 is not authenticated")
    if build_receipt.get("exit_code") != 0:
        raise QualificationError(f"{receipt_path}.exit_code must be zero")
    if build_receipt.get("clean_before") is not True or build_receipt.get("clean_after") is not True:
        raise QualificationError(f"{receipt_path} does not prove a clean build tree")
    if build_receipt.get("source_manifest_verified_before_build") is not True or build_receipt.get("source_manifest_verified_after_build") is not True:
        raise QualificationError(f"{receipt_path} does not prove source manifest verification before/after build")
    source_manifest_sha = _sha(
        build_receipt.get("source_manifest_sha256"), f"{receipt_path}.source_manifest_sha256"
    )
    source_manifest_files = build_receipt.get("source_manifest_files")
    if isinstance(source_manifest_files, bool) or not isinstance(source_manifest_files, int) or source_manifest_files <= 0:
        raise QualificationError(f"{receipt_path}.source_manifest_files must be positive")
    source_manifest_name = _string(
        build_receipt.get("source_manifest"), f"{receipt_path}.source_manifest"
    )
    source_bindings_path = path.parent / "source-bindings.json"
    source_bindings = _object(normal.load_json(source_bindings_path), str(source_bindings_path))
    source_roles = _object(source_bindings.get("roles"), f"{source_bindings_path}.roles")
    source_role = _object(source_roles.get(role), f"{source_bindings_path}.roles.{role}")
    source_manifest_path = path.parent / _string(source_role.get("path"), f"{source_bindings_path}.roles.{role}.path")
    if source_manifest_path.name != source_manifest_name:
        raise QualificationError(f"{receipt_path}.source_manifest does not match source-bindings.json")
    if normal.raw_sha256(source_manifest_path) != source_manifest_sha:
        raise QualificationError(f"{receipt_path}.source_manifest_sha256 does not match its retained manifest")
    if source_role.get("revision") != revision or source_role.get("sha256") != source_manifest_sha or source_role.get("files") != source_manifest_files:
        raise QualificationError(f"{receipt_path} source manifest binding does not match its role")
    manifest_value = _object(normal.load_json(source_manifest_path), str(source_manifest_path))
    if len(manifest_value) != source_manifest_files:
        raise QualificationError(f"{source_manifest_path} file count differs from its binding")
    argv = build_receipt.get("argv")
    if not isinstance(argv, list) or argv[:2] != ["cargo", "build"]:
        raise QualificationError(f"{receipt_path}.argv is not an authenticated cargo build")
    build_cwd = _string(build_receipt.get("cwd"), f"{receipt_path}.cwd")
    if binding.get("clean_build") is not True:
        raise QualificationError(f"{path}.clean_build must be true")

    source_binding: dict[str, Any] = {
        "kind": "clean_build_receipt_revision",
        "revision": revision,
        "build_receipt_path": receipt_name,
        "build_receipt_sha256": receipt_digest,
        "build_receipt_canonical_sha256": normal.canonical_sha256(build_receipt),
        "build_cwd": build_cwd,
        "source_manifest": source_manifest_name,
        "source_manifest_sha256": source_manifest_sha,
        "source_manifest_files": source_manifest_files,
    }
    optional_source = binding.get("source")
    if optional_source is not None:
        source_object = _object(optional_source, f"{path}.source")
        source_name = _string(source_object.get("path"), f"{path}.source.path")
        source_path = Path(source_name)
        if source_path.is_absolute() or ".." in source_path.parts:
            raise QualificationError(f"{path}.source.path must be bundle-relative")
        source_path = path.parent / source_path
        source_digest = _sha(source_object.get("sha256"), f"{path}.source.sha256")
        if normal.raw_sha256(source_path) != source_digest:
            raise QualificationError(f"{path}.source.sha256 does not match its manifest")
        files = source_object.get("files")
        if isinstance(files, bool) or not isinstance(files, int) or files <= 0:
            raise QualificationError(f"{path}.source.files must be positive")
        source_binding.update(
            {"kind": "explicit_source_manifest", "path": source_name, "sha256": source_digest, "files": files}
        )
    if binding.get("source_manifest_sha256") is not None:
        binding_source_sha = _sha(
            binding.get("source_manifest_sha256"), f"{path}.source_manifest_sha256"
        )
        if binding_source_sha != source_manifest_sha:
            raise QualificationError(f"{path}.source_manifest_sha256 does not match its build receipt")
    return path, {
        "revision": revision,
        "binary_sha256": binary,
        "bytes": binary_bytes,
        "build_receipt": receipt_name,
        "build_receipt_sha256": receipt_digest,
        "build_cwd": build_cwd,
        "source_manifest": source_manifest_name,
        "source_manifest_sha256": source_manifest_sha,
        "source_manifest_files": source_manifest_files,
        "source_binding": source_binding,
        "binding_sha256": normal.raw_sha256(path),
    }


def _fixed_receipt(
    path: Path,
    lane: str,
    binding_path: Path,
    binding: Mapping[str, Any],
    *,
    role: str,
    samples: int,
    warmups: int,
    full: bool = False,
) -> dict[str, Any]:
    receipt = _object(normal.load_json(path), str(path))
    if receipt.get("schema") != "litchi-0467-capture-fixed-v1":
        raise QualificationError(f"{path}.schema is not the fixed capture schema")
    if receipt.get("lane") != lane or receipt.get("role") != role:
        raise QualificationError(f"{path} lane/role identity mismatch")
    if receipt.get("revision") != binding["revision"]:
        raise QualificationError(f"{path}.revision does not match its binding")
    if receipt.get("binary_sha256") != binding["binary_sha256"]:
        raise QualificationError(f"{path}.binary_sha256 does not match its binding")
    if receipt.get("binding_sha256") != normal.raw_sha256(binding_path):
        raise QualificationError(f"{path}.binding_sha256 does not match its binding file")
    if receipt.get("clean_before") is not True or receipt.get("clean_after") is not True:
        raise QualificationError(f"{path} does not prove a clean checkout before/after capture")
    if receipt.get("binary_unchanged") is not True or receipt.get("report_metadata_matches_clean_role") is not True:
        raise QualificationError(f"{path} does not prove binary and report identity")
    if receipt.get("samples") != samples or receipt.get("warmups") != warmups:
        raise QualificationError(f"{path} sample/warmup counts do not match its lane")
    argv = receipt.get("argv")
    if not isinstance(argv, list) or argv[:3] != ["taskset", "-c", CPU]:
        raise QualificationError(f"{path}.argv does not pin the fixed capture to CPU 2")
    if "heaptrack" in argv:
        raise QualificationError(f"{path}.argv must not use Heaptrack for fixed normal lanes")
    if "--workers" not in argv or argv[argv.index("--workers") + 1] != str(WORKERS):
        raise QualificationError(f"{path}.argv does not use one worker")
    if receipt.get("exit_code") != 0:
        raise QualificationError(f"{path}.exit_code must be zero")
    checkout_argv = receipt.get("checkout_argv")
    if not isinstance(checkout_argv, list) or checkout_argv[:3] != ["git", "checkout", "--detach"]:
        raise QualificationError(f"{path}.checkout_argv does not prove a detached revision checkout")
    if binding["revision"] not in checkout_argv:
        raise QualificationError(f"{path}.checkout_argv does not name its bound revision")
    environment = _object(receipt.get("environment"), f"{path}.environment")
    expected_environment = {
        "RUSTUP_TOOLCHAIN": "1.98.1",
        "RUSTFLAGS": "-C force-frame-pointers=yes -C force-unwind-tables=yes",
        "CARGO_PROFILE_RELEASE_DEBUG": "1",
        "CARGO_INCREMENTAL": "0",
        "CARGO_BUILD_JOBS": "4",
        "DEBUGINFOD_URLS": "",
        "LC_ALL": "C",
    }
    if environment != expected_environment:
        raise QualificationError(f"{path}.environment differs from the frozen fixed capture environment")
    if full:
        if receipt.get("cases") is not None or receipt.get("xlsx_shape") is not None:
            raise QualificationError(f"{path} full guard receipt must not carry normal workload selectors")
    else:
        if receipt.get("cases") != list(FIXED_CASES):
            raise QualificationError(f"{path}.cases does not select the fixed four-row workload")
        if receipt.get("xlsx_shape") != XLSX_SHAPE:
            raise QualificationError(f"{path}.xlsx_shape does not select dense-wide")
    _string(receipt.get("cwd"), f"{path}.cwd")
    return receipt


def _fixed_report(
    report: Mapping[str, Any],
    label: str,
    binding: Mapping[str, Any],
) -> tuple[dict[str, Any], dict[tuple[str, str], dict[str, Any]], dict[str, Any]]:
    normal._report_profile(report, label)
    environment = _object(report.get("environment"), f"{label}.environment")
    if environment.get("git_worktree_dirty") is not False:
        raise QualificationError(f"{label}.environment.git_worktree_dirty must be false")
    if environment.get("cpu_affinity") != CPU:
        raise QualificationError(f"{label}.environment.cpu_affinity must be CPU 2")
    if _revision(environment.get("git_revision"), f"{label}.environment.git_revision") != binding["revision"]:
        raise QualificationError(f"{label}.environment.git_revision does not match its binding")
    configuration = _object(report.get("configuration"), f"{label}.configuration")
    if configuration.get("samples_per_case") != SAMPLES:
        raise QualificationError(f"{label}.configuration.samples_per_case must be {SAMPLES}")
    if configuration.get("warmup_iterations_per_case") != WARMUPS:
        raise QualificationError(f"{label}.configuration.warmup_iterations_per_case must be {WARMUPS}")
    _exact_list(configuration.get("cases"), FIXED_CASES, f"{label}.configuration.cases")
    _exact_list(configuration.get("writer_shapes"), ("tiny", "large", "payload-heavy"), f"{label}.configuration.writer_shapes")
    _exact_list(configuration.get("xlsx_shapes"), (XLSX_SHAPE,), f"{label}.configuration.xlsx_shapes")
    _exact_list(configuration.get("execution_workers"), (WORKERS,), f"{label}.configuration.execution_workers")
    binary = _object(report.get("binary_identity"), f"{label}.binary_identity")
    if _sha(binary.get("binary_sha256"), f"{label}.binary_identity.binary_sha256") != binding["binary_sha256"]:
        raise QualificationError(f"{label}.binary_identity.binary_sha256 does not match its binding")
    if binary.get("binary_bytes") != binding["bytes"]:
        raise QualificationError(f"{label}.binary_identity.binary_bytes does not match its binding")
    results = report.get("results")
    if not isinstance(results, list) or len(results) != len(EXPECTED_ROW_SHAPES):
        raise QualificationError(f"{label}.results must contain exactly {len(EXPECTED_ROW_SHAPES)} rows")
    indexed: dict[tuple[str, str], dict[str, Any]] = {}
    observed_shapes: set[tuple[str, str]] = set()
    for index, raw_row in enumerate(results):
        row = _object(raw_row, f"{label}.results[{index}]")
        case = _string(row.get("case"), f"{label}.results[{index}].case")
        corpus = _object(row.get("corpus"), f"{label}.results[{index}].corpus")
        shape = _string(corpus.get("shape"), f"{label}.results[{index}].corpus.shape")
        identity = (case, shape)
        if identity not in EXPECTED_ROW_SHAPES:
            raise QualificationError(f"{label}.results[{index}] is outside the fixed four-row scope")
        if identity in observed_shapes:
            raise QualificationError(f"{label} contains duplicate {case}/{shape}")
        if case == DOC_CASE and corpus.get("name") != f"doc-{shape}":
            raise QualificationError(f"{label}.{case}[{shape}] corpus name is not doc-{shape}")
        if case == XLSX_CASE and corpus.get("name") != "xlsx-dense-wide":
            raise QualificationError(f"{label}.{case}[{shape}] corpus name is not xlsx-dense-wide")
        elapsed = _object(row.get("elapsed_ns"), f"{label}.{case}[{shape}].elapsed_ns")
        if not isinstance(elapsed.get("samples"), list) or len(elapsed["samples"]) != SAMPLES:
            raise QualificationError(f"{label}.{case}[{shape}] must contain {SAMPLES} samples")
        try:
            statistics = perf_abba_summary.recompute_statistics(
                elapsed, f"{label}.{case}[{shape}].elapsed_ns"
            )
        except Exception as error:
            raise QualificationError(f"{label}.{case}[{shape}] statistic validation failed: {error}") from error
        key = (case, normal.canonical_sha256(corpus))
        if key in indexed:
            raise QualificationError(f"{label} contains duplicate corpus identity for {case}/{shape}")
        indexed[key] = {"row": row, "corpus": corpus, "statistics": statistics}
        observed_shapes.add(identity)
    if observed_shapes != EXPECTED_ROW_SHAPES:
        raise QualificationError(f"{label} row scope differs: {sorted(observed_shapes)!r}")
    return dict(report), indexed, configuration


def _fixed_catalog(root: Path, lane_dir: Path, report: Mapping[str, Any], label: str) -> dict[str, Any]:
    try:
        return normal._catalog_file_identity(root, lane_dir, report, label)
    except normal.AnalysisError as error:
        raise QualificationError(str(error)) from error


def _load_fixed_normal(
    root: Path,
    lane: str,
    bindings: Mapping[str, tuple[Path, Mapping[str, Any]]],
) -> dict[str, Any]:
    lane_dir = root / lane
    if not lane_dir.is_dir():
        raise QualificationError(f"missing fixed lane directory {lane_dir}")
    role = FIXED_ROLES[lane]
    binding_path, binding = bindings[role]
    receipt = _fixed_receipt(
        lane_dir / "receipt.json",
        lane,
        binding_path,
        binding,
        role=role,
        samples=SAMPLES,
        warmups=WARMUPS,
    )
    report_path = lane_dir / "report.json"
    report, indexed, configuration = _fixed_report(
        _object(normal.load_json(report_path), str(report_path)), lane, binding
    )
    resource = normal._resource_rss(lane_dir / "resource.log")
    catalog = _fixed_catalog(root, lane_dir, report, lane)
    return {
        "lane": lane,
        "role": role,
        "path": lane_dir,
        "binding_path": binding_path,
        "binding": dict(binding),
        "receipt": receipt,
        "report": report,
        "indexed": indexed,
        "configuration": configuration,
        "resource": resource,
        "catalog": catalog,
    }


def _check_fixed_checkout_paths(records: Sequence[Mapping[str, Any]]) -> dict[str, Any]:
    capture_paths = {_string(record["receipt"].get("cwd"), f"{record['lane']}.receipt.cwd") for record in records}
    build_paths = {record["binding"]["build_cwd"] for record in records}
    if len(capture_paths) != 1:
        raise QualificationError(f"fixed capture cwd differs between lanes: {sorted(capture_paths)!r}")
    if len(build_paths) != 1:
        raise QualificationError(f"fixed build cwd differs between lanes: {sorted(build_paths)!r}")
    capture_path = next(iter(capture_paths))
    build_path = next(iter(build_paths))
    if capture_path != build_path:
        raise QualificationError(
            f"fixed build/capture cwd differs: build={build_path!r}, capture={capture_path!r}"
        )
    return {"path": capture_path, "build_cwd": build_path, "capture_cwd": capture_path, "same_path": True}


def _check_identity(records: Mapping[str, Mapping[str, Any]]) -> None:
    control = records["a1"]
    candidate = records["b1"]
    if control["binding"]["revision"] == candidate["binding"]["revision"]:
        raise QualificationError("fixed control and candidate revisions must differ")
    if control["binding"]["binary_sha256"] == candidate["binding"]["binary_sha256"]:
        raise QualificationError("fixed control and candidate binaries must differ")
    for role, first, second in (("control", "a1", "a2"), ("candidate", "b1", "b2")):
        if records[first]["binding"]["revision"] != records[second]["binding"]["revision"]:
            raise QualificationError(f"fixed {role} revision differs between ABBA legs")
        if records[first]["binding"]["binary_sha256"] != records[second]["binding"]["binary_sha256"]:
            raise QualificationError(f"fixed {role} binary differs between ABBA legs")
    catalog_ids = {record["catalog"]["corpus_identity_sha256"] for record in records.values()}
    if len(catalog_ids) != 1:
        raise QualificationError("fixed lanes do not retain one corpus catalog identity")
    if records["a1"]["catalog"]["canonical_sha256"] == records["b1"]["catalog"]["canonical_sha256"]:
        raise QualificationError("fixed control/candidate catalog custody digests must differ")
    row_keys = [set(record["indexed"]) for record in records.values()]
    if any(keys != row_keys[0] for keys in row_keys[1:]):
        raise QualificationError("fixed ABBA row corpus identities differ")
    for key in sorted(row_keys[0]):
        corpus_identity = normal.canonical_sha256(records["a1"]["indexed"][key]["corpus"])
        if any(normal.canonical_sha256(record["indexed"][key]["corpus"]) != corpus_identity for record in records.values()):
            raise QualificationError(f"fixed corpus identity differs for {key[0]}")


def _row_results(records: Mapping[str, Mapping[str, Any]]) -> list[dict[str, Any]]:
    indexed = {leg: records[leg]["indexed"] for leg in FIXED_LEGS}
    rss = {leg: records[leg]["resource"] for leg in FIXED_LEGS}
    lower_indexed = {leg: indexed[leg] for leg in FIXED_LEGS}
    results = [normal._row_result(key, lower_indexed, rss) for key in sorted(indexed["a1"])]
    for result in results:
        result["claim_role"] = "primary" if result["case"] == XLSX_CASE else "guard_only"
    return results


def _summary_result(summary: Mapping[str, Any], case: str, shape: str) -> dict[str, Any]:
    results = summary.get("results")
    if not isinstance(results, list):
        raise QualificationError("canonical ABBA summary.results must be a list")
    matches = [item for item in results if isinstance(item, dict) and item.get("case") == case and item.get("shape") == shape]
    if len(matches) != 1:
        raise QualificationError(f"canonical ABBA summary has {len(matches)} rows for {case}/{shape}")
    return _object(matches[0], f"summary.{case}[{shape}]")


def _cell_list(summary_result: Mapping[str, Any], field: str) -> list[dict[str, str]]:
    elapsed = _object(summary_result.get("elapsed_ns"), "summary result.elapsed_ns")
    values = elapsed.get(field, [])
    if not isinstance(values, list):
        raise QualificationError(f"summary result.elapsed_ns.{field} must be a list")
    corpus = _object(summary_result.get("corpus"), "summary result.corpus")
    name = _string(corpus.get("name"), "summary result.corpus.name")
    for statistic in values:
        if statistic not in STATISTICS:
            raise QualificationError(f"summary result.elapsed_ns.{field} contains {statistic!r}")
    return [{"case": XLSX_CASE, "corpus": name, "statistic": statistic} for statistic in values]


def _claim_proposal(summary_result: Mapping[str, Any]) -> dict[str, Any]:
    accepted = _cell_list(summary_result, "accepted_statistics")
    adverse = _cell_list(summary_result, "adverse_both_statistics")
    corpus = _object(summary_result.get("corpus"), "summary result.corpus")
    scope_corpus = {
        key: corpus[key]
        for key in ("name", "archive_sha256", "generator", "shape", "package_format")
        if key in corpus
    }
    reasons = ["accepted_primary_latency_cells"] if accepted else ["no_accepted_primary_latency_cells"]
    if adverse:
        reasons.append("explicit_adverse_primary_cells")
    entry: dict[str, Any] = {
        "id": CLAIM_ID,
        "change_id": PACKAGE_CHANGE_ID,
        "claim_class": "optimization",
        "status": "landed" if accepted else "rejected",
        "code_state": "landed" if accepted else "not_landed",
        "reason_codes": reasons,
        "scope": {"format": "xlsx", "selectors": [XLSX_CASE], "corpora": [scope_corpus]},
        "latency_evidence": {
            "evidence_id": EVIDENCE_ID,
            "allowed_statistics": list(STATISTICS),
            "accepted_cells": len(accepted),
            "adverse_both_cells": len(adverse),
            "accepted_statistics": accepted,
            "adverse_both_statistics": adverse,
            "doc_rows_are_guards_only": True,
        },
        "documentation": [DOCUMENTATION],
    }
    return {"proposal_only": True, "entry": entry}


def _load_full_fixed_guard(
    root: Path,
    lane: str,
    bindings: Mapping[str, tuple[Path, Mapping[str, Any]]],
) -> dict[str, Any]:
    lane_dir = root / lane
    if not lane_dir.is_dir():
        raise QualificationError(f"missing fixed full guard lane {lane_dir}")
    role = "control" if lane.startswith("A-") else "candidate"
    binding_path, binding = bindings[role]
    receipt = _fixed_receipt(
        lane_dir / "receipt.json", lane, binding_path, binding, role=role,
        samples=FULL_SAMPLES, warmups=FULL_WARMUPS,
        full=True,
    )
    report_path = lane_dir / "report.json"
    report = _object(normal.load_json(report_path), str(report_path))
    try:
        indexed = normal._validate_guard_report(
            report, lane, binding, samples=FULL_SAMPLES, warmups=FULL_WARMUPS,
            expected_rows=FULL_ROWS,
        )
    except normal.AnalysisError as error:
        raise QualificationError(str(error)) from error
    resource = normal._resource_rss(lane_dir / "resource.log")
    catalog = _fixed_catalog(root, lane_dir, report, lane)
    return {
        "lane": lane, "role": role, "path": lane_dir, "binding_path": binding_path,
        "binding": dict(binding), "receipt": receipt, "report": report,
        "indexed": indexed, "resource": resource, "catalog": catalog,
    }


def _full_guard(root: Path, bindings: Mapping[str, tuple[Path, Mapping[str, Any]]], protocol: Mapping[str, Any]) -> dict[str, Any]:
    selected_pair: tuple[str, str] | None = None
    for pair in FULL_FIXED_LANE_PAIRS:
        present = [(root / lane).is_dir() for lane in pair]
        if any(present) and not all(present):
            raise QualificationError(f"fixed full guard pair is incomplete: {pair!r}")
        if all(present):
            selected_pair = pair
            break
    if selected_pair is None:
        raise QualificationError("fixed full guard lanes are missing")
    control = _load_full_fixed_guard(root, selected_pair[0], bindings)
    candidate = _load_full_fixed_guard(root, selected_pair[1], bindings)
    checkout = _check_fixed_checkout_paths((control, candidate))
    if control["catalog"]["corpus_identity_sha256"] != candidate["catalog"]["corpus_identity_sha256"]:
        raise QualificationError("fixed full guard corpus catalogs differ")
    if control["catalog"]["canonical_sha256"] == candidate["catalog"]["canonical_sha256"]:
        raise QualificationError("fixed full guard control/candidate catalog custody digests must differ")
    try:
        comparison = perf_compare.compare_reports(
            normal._comparison_report(control["report"]),
            normal._comparison_report(candidate["report"]),
            _object(normal.load_json(REPO_ROOT / "docs/performance/perf-regression-policy-v1.json"), "policy"),
        )
    except Exception as error:
        raise QualificationError(f"fixed full guard comparison failed: {error}") from error
    review = normal._full_guard_latency_review(control, candidate)
    rss_pair = normal._ratio_percent(
        control["resource"]["max_rss_kib"], candidate["resource"]["max_rss_kib"],
        "fixed_full_guard.process_rss",
    )
    return {
        "lanes": [selected_pair[0], selected_pair[1]],
        "scope": "fixed checkout full default matrix, 201 rows, 15 samples/3 warmups",
        "fixed_checkout": checkout,
        "control": {"binding": control["binding"], "catalog": control["catalog"], "resource": control["resource"]},
        "candidate": {"binding": candidate["binding"], "catalog": candidate["catalog"], "resource": candidate["resource"]},
        "comparison": comparison,
        "latency_review_5_percent": review,
        "process_rss": {
            "scope": "whole process per fixed full guard lane",
            "control_kib": control["resource"]["max_rss_kib"],
            "candidate_kib": candidate["resource"]["max_rss_kib"],
            "pair": rss_pair,
            "review_triggered": rss_pair["regression_over_5_percent"],
        },
        "verification": {
            "fixed_checkout_same_path_verified": True,
            "clean_worktrees_verified": True,
            "distinct_revisions_verified": control["binding"]["revision"] != candidate["binding"]["revision"],
            "distinct_binaries_verified": control["binding"]["binary_sha256"] != candidate["binding"]["binary_sha256"],
            "catalog_identity_verified": True,
            "default_policy_comparison_verified": True,
            "instrumented_latency_not_compared": True,
        },
    }


def qualify(root: Path = ROOT, *, protocol_path: Path | None = None) -> dict[str, Any]:
    root = root.resolve()
    protocol_file = _protocol_path(root, protocol_path)
    protocol = _validate_protocol(protocol_file)
    bindings = {role: _fixed_binding(root, role) for role in ("control", "candidate")}
    records = {
        leg: _load_fixed_normal(root, lane, bindings)
        for lane, leg in zip(FIXED_LANES, FIXED_LEGS)
    }
    _check_identity(records)
    checkout = _check_fixed_checkout_paths(list(records.values()))
    configurations = {normal.canonical_sha256(record["configuration"]) for record in records.values()}
    if len(configurations) != 1:
        raise QualificationError("fixed report configurations differ between ABBA lanes")
    try:
        summary = perf_abba_summary.summarize_reports(
            [records[leg]["report"] for leg in FIXED_LEGS],
            drift_ceilings=DRIFT_CEILINGS,
            cases=FIXED_CASES,
        )
    except Exception as error:
        raise QualificationError(f"fixed canonical ABBA summary validation failed: {error}") from error
    summary_results = summary.get("results")
    if not isinstance(summary_results, list) or len(summary_results) != len(EXPECTED_ROW_SHAPES):
        raise QualificationError("fixed canonical ABBA summary must contain exactly four rows")
    primary_summary = _summary_result(summary, XLSX_CASE, XLSX_SHAPE)
    rows = _row_results(records)
    primary_row = next(row for row in rows if row["case"] == XLSX_CASE)
    doc_rows = [row for row in rows if row["case"] == DOC_CASE]
    if len(doc_rows) != len(DOC_SHAPES):
        raise QualificationError("fixed qualification must retain all three DOC guard rows")
    full_guard = _full_guard(root, bindings, protocol)
    lane_records: dict[str, Any] = {}
    for lane, leg in zip(FIXED_LANES, FIXED_LEGS):
        record = records[leg]
        report_path = record["path"] / "report.json"
        receipt_path = record["path"] / "receipt.json"
        lane_records[lane] = {
            "leg": leg,
            "role": record["role"],
            "path": str(record["path"].relative_to(root)),
            "binding_path": str(record["binding_path"].relative_to(root)),
            "binding": record["binding"],
            "report": {
                "path": str(report_path.relative_to(root)),
                "sha256": normal.raw_sha256(report_path),
                "canonical_sha256": normal.canonical_sha256(record["report"]),
                "bytes": report_path.stat().st_size,
            },
            "receipt": {
                "path": str(receipt_path.relative_to(root)),
                "sha256": normal.raw_sha256(receipt_path),
                "bytes": receipt_path.stat().st_size,
            },
            "catalog": record["catalog"],
            "resource": record["resource"],
        }
    accepted = _cell_list(primary_summary, "accepted_statistics")
    adverse = _cell_list(primary_summary, "adverse_both_statistics")
    return {
        "schema": SCHEMA,
        "purpose": "fixed-checkout 500-sample ABBA qualification with XLSX primary and DOC guard rows",
        "protocol": {
            "path": str(protocol_file.relative_to(root)) if protocol_file.is_relative_to(root) else str(protocol_file),
            "sha256": normal.raw_sha256(protocol_file),
            "canonical_sha256": normal.canonical_sha256(protocol),
            "schema": protocol.get("schema"),
            "order": list(FIXED_LANES),
            "samples": SAMPLES,
            "warmups": WARMUPS,
            "cpu": 2,
            "workers": WORKERS,
        },
        "fixed_checkout": checkout,
        "scope": {
            "primary_case": XLSX_CASE,
            "primary_shape": XLSX_SHAPE,
            "doc_guard_case": DOC_CASE,
            "doc_guard_shapes": list(DOC_SHAPES),
            "row_count": len(EXPECTED_ROW_SHAPES),
            "configuration": records["a1"]["configuration"],
        },
        "bindings": {
            role: {**dict(binding), "path": str(path.relative_to(root))}
            for role, (path, binding) in bindings.items()
        },
        "lanes": lane_records,
        "rows": rows,
        "primary": {"summary": primary_summary, "row": primary_row},
        "doc_guard": {
            "purpose": "retain all three DOC rows as deliberate guards; never register them as XLSX claim scope",
            "rows": doc_rows,
            "drift_policy_percent": dict(DRIFT_CEILINGS),
        },
        "process_rss": {
            "scope": "whole process per fixed normal lane; GNU time maximum resident set size",
            "kib": {leg: records[leg]["resource"]["max_rss_kib"] for leg in FIXED_LEGS},
            "pairings": {
                "a1_control_to_b1_candidate": normal._ratio_percent(records["a1"]["resource"]["max_rss_kib"], records["b1"]["resource"]["max_rss_kib"], "fixed.process_rss.a1-b1"),
                "a2_control_to_b2_candidate": normal._ratio_percent(records["a2"]["resource"]["max_rss_kib"], records["b2"]["resource"]["max_rss_kib"], "fixed.process_rss.a2-b2"),
            },
            "same_implementation_drift": {
                "control": normal._drift(records["a1"]["resource"]["max_rss_kib"], records["a2"]["resource"]["max_rss_kib"], "fixed.process_rss.control"),
                "candidate": normal._drift(records["b1"]["resource"]["max_rss_kib"], records["b2"]["resource"]["max_rss_kib"], "fixed.process_rss.candidate"),
            },
            "review_threshold_percent": REVIEW_THRESHOLD_PERCENT,
            "claim_role": "descriptive process guard only",
        },
        "abba_summary": summary,
        "full_guard": full_guard,
        "claim_registration": {
            "registry_path": "docs/performance/claim-registry-v1.json",
            "latency_policy": "latency-abba-v1",
            "minimum_samples_per_case": SAMPLES,
            "drift_ceiling_percent": dict(DRIFT_CEILINGS),
            "accepted_primary_cells": accepted,
            "adverse_primary_cells": adverse,
            "strict_latency_claim_eligible": bool(accepted),
            "registry_entry_ready": False,
            "reason": "proposal only; review individual outcomes and package custody before registration",
            "doc_rows_are_guards_only": True,
            "registry_entry_proposal": _claim_proposal(primary_summary),
        },
        "verification": {
            "fixed_checkout_same_path_verified": True,
            "clean_worktrees_verified": True,
            "report_clean_metadata_verified": True,
            "distinct_revisions_verified": True,
            "distinct_binaries_verified": True,
            "build_receipt_bindings_verified": True,
            "source_binding_verified_via_build_receipts": True,
            "configuration_identity_verified": True,
            "row_identity_verified": True,
            "catalog_identity_verified": True,
            "statistics_recomputed_from_500_samples": True,
            "canonical_abba_summary_verified": True,
            "drift_policy_verified": True,
            "doc_guard_rows_retained": True,
            "process_rss_receipts_verified": True,
            "process_rss_review_trigger_is_not_claim": True,
            "full_guard_verified": True,
            "instrumented_latency_not_used": True,
        },
    }


def _write_json(path: Path, value: Mapping[str, Any]) -> None:
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        with path.open("x", encoding="utf-8") as stream:
            json.dump(value, stream, indent=2, sort_keys=True, allow_nan=False)
            stream.write("\n")
    except FileExistsError as error:
        raise QualificationError(f"output already exists: {path}") from error
    except OSError as error:
        raise QualificationError(f"cannot write {path}: {error}") from error


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, default=ROOT)
    parser.add_argument("--protocol", type=Path)
    parser.add_argument("--json-out", type=Path)
    parser.add_argument("--summary-out", type=Path)
    return parser


def main(argv: Iterable[str] | None = None) -> int:
    args = build_parser().parse_args(list(argv) if argv is not None else None)
    try:
        root = args.root.resolve()
        result = qualify(root, protocol_path=args.protocol.resolve() if args.protocol else None)
        if args.json_out is not None:
            _write_json(args.json_out, result)
        if args.summary_out is not None:
            _write_json(args.summary_out, result["abba_summary"])
        json.dump(result, sys.stdout, indent=2, sort_keys=True, allow_nan=False)
        sys.stdout.write("\n")
        return 0
    except (QualificationError, OSError, ValueError) as error:
        print(f"{SCHEMA}: {error}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
