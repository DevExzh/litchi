#!/usr/bin/env python3
"""Fail-closed replay checks for the 0511 CFB/OLE2 XLS capture.

This verifier only consumes the receipts, reports, catalogs, logs, and frozen
source manifests produced by ``capture.py``.  It never rebuilds the harness,
reruns a measurement, or treats missing evidence as a zero.  The normal XLS
and CFB lanes are the only latency comparison; the allocator lane is checked
for aligned allocation vectors but its elapsed values are deliberately kept
out of the native comparison.  The callgrind lane is a five-sample scoped
diagnostic for ``SourceBackedWorkbook::from_read_at_with_limits`` and is also
excluded from native timing comparisons.
"""

from __future__ import annotations

import copy
import datetime as _datetime
import hashlib
import json
import math
import re
import subprocess
import sys
from pathlib import Path
from typing import Any

HERE = Path(__file__).resolve().parent
REPO = HERE.parents[3]
sys.path.insert(0, str(REPO))

from tools.summarize_crud_baseline import _validate_elapsed
from tools.validate_perf_corpus_binding import validate_binding


BASE_REVISION = "b0b9ebb577d616a9a0ed93f16619a1b79f3b3bc7"
FAT_SOURCE = "crates/litchi-cfb/src/file.rs"
SHA256_RE = re.compile(r"^[0-9a-f]{64}$")
RSS_RE = re.compile(r"Maximum resident set size \(kbytes\):\s*(\d+)")
CALLGRIND_SUMMARY_RE = re.compile(r"^summary:\s*(\d+)\s*$", re.MULTILINE)

XLS_CASES = (
    "xls_semantic_open",
    "xls_eager_open_list_worksheets",
    "xls_eager_open_one_cell",
    "xls_source_backed_open",
    "xls_source_backed_open_list_worksheets",
    "xls_source_backed_open_one_cell",
    "xls_owned_source_open",
    "xls_owned_source_open_list_worksheets",
    "xls_owned_source_open_one_cell",
)
XLS_CASE_SET = set(XLS_CASES)
CFB_CASES = ("cfb_open",)
SOURCE_XLS_CASES = {
    "xls_source_backed_open",
    "xls_source_backed_open_list_worksheets",
    "xls_source_backed_open_one_cell",
}
CALLGRIND_CASE = "xls_owned_source_open_one_cell"
ALLOCATOR_VECTOR_FIELDS = (
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
ELAPSED_METRICS = ("p50", "p95", "p99", "mean")
DRIFT_LIMITS = {"p50": 5.0, "mean": 5.0, "p95": 10.0, "p99": 15.0}


class VerificationError(ValueError):
    """A missing, malformed, or inconsistent evidence item."""


def fail(message: str) -> None:
    raise VerificationError(message)


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def load(path: Path) -> Any:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        fail(f"cannot load {path}: {error}")


def sha(path: Path) -> str:
    try:
        return hashlib.sha256(path.read_bytes()).hexdigest()
    except OSError as error:
        fail(f"cannot hash {path}: {error}")


def sha256_text(value: str) -> str:
    return hashlib.sha256(value.encode("utf-8")).hexdigest()


def check_hash(value: Any, context: str) -> str:
    require(isinstance(value, str) and SHA256_RE.fullmatch(value) is not None,
            f"{context} is not a lowercase SHA-256")
    return value


def canonical(value: Any) -> str:
    try:
        return json.dumps(value, sort_keys=True, separators=(",", ":"), allow_nan=False)
    except (TypeError, ValueError) as error:
        fail(f"value is not canonical JSON: {error}")


def stage_dir(stage: str) -> Path:
    require(stage in {"before", "after"}, f"unknown stage {stage!r}")
    return HERE / stage


def manifest_path(stage: str) -> Path:
    # run.py writes the preserved control manifest under before/ and the
    # candidate manifest at the capture root, which is also the after build
    # directory used by capture.py.
    return (HERE / "before" / "source-manifest.json") if stage == "before" else HERE / "source-manifest.json"


def build_receipt_path(stage: str, allocator: bool = False) -> Path:
    if allocator:
        return stage_dir(stage) / "allocator-build-receipt.json"
    return (HERE / "before" / "build-receipt.json") if stage == "before" else HERE / "build-receipt.json"


def source_manifest(stage: str) -> dict[str, str]:
    value = load(manifest_path(stage))
    require(isinstance(value, dict), f"{stage} source manifest must be an object")
    for path, digest in value.items():
        require(isinstance(path, str) and path and not Path(path).is_absolute(),
                f"{stage} source manifest has an invalid path {path!r}")
        check_hash(digest, f"{stage} source manifest {path}")
    return value


def repository_file(relative: str) -> Path:
    path = (REPO / relative).resolve()
    try:
        path.relative_to(REPO.resolve())
    except ValueError:
        fail(f"repository path escapes checkout: {relative!r}")
    require(path.is_file(), f"repository source file is missing: {relative}")
    return path


def git_blob_bytes(revision: str, relative: str) -> bytes:
    try:
        data = subprocess.check_output(
            ["git", "show", f"{revision}:{relative}"], cwd=REPO
        )
    except (OSError, subprocess.CalledProcessError) as error:
        fail(f"cannot read {revision}:{relative} from git: {error}")
    return data


def git_blob_sha(revision: str, relative: str) -> str:
    return hashlib.sha256(git_blob_bytes(revision, relative)).hexdigest()


def verify_fat_source_change(base_revision: str) -> None:
    """Keep the candidate bounded to the already-reserved FAT-entry loop.

    The candidate may use a direct push or an equivalent exact-sized batch
    insertion.  The verifier therefore checks the unchanged context and the
    safety-relevant facts instead of freezing one implementation spelling or
    an intermediate binary/profile result.
    """

    before_bytes = git_blob_bytes(base_revision, FAT_SOURCE)
    after_path = repository_file(FAT_SOURCE)
    try:
        before_text = before_bytes.decode("utf-8")
        after_text = after_path.read_text(encoding="utf-8")
    except (UnicodeError, OSError) as error:
        fail(f"cannot decode FAT source for bounded diff: {error}")

    loop_start = "        for &sector_id in &fat_sectors {\n"
    loop_end = "\n        for sector in fat_sectors {\n"

    def split_loop(text: str) -> tuple[str, str, str]:
        start = text.find(loop_start)
        require(start >= 0, "FAT sector loop is missing from the source")
        body_start = start + len(loop_start)
        end = text.find(loop_end, body_start)
        require(end >= 0, "FAT marker-validation loop is missing from the source")
        return text[:start], text[body_start:end], text[end:]

    before_prefix, before_body, before_suffix = split_loop(before_text)
    after_prefix, after_body, after_suffix = split_loop(after_text)
    require(before_prefix == after_prefix and before_suffix == after_suffix,
            "candidate changes extend outside the FAT-entry loop")
    require(before_body != after_body, "candidate FAT-entry loop is unchanged")
    reservation = 'let mut fat = try_vec_with_capacity(fat_entry_count, "FAT entries")?;'
    require(reservation in after_prefix,
            "candidate FAT loop no longer follows the exact fallible reservation")
    require("self.read_sector_into(sector_id, sector_data)?;" in after_body,
            "candidate FAT loop no longer reads each declared sector")
    decodes_entries = (
        ("read_u32_le" in after_body and '"FAT entry"' in after_body)
        or "u32::from_le_bytes" in after_body
    )
    require(decodes_entries, "candidate FAT loop no longer decodes FAT entries")
    require("try_push(&mut fat" not in after_body,
            "candidate retains the per-entry fallible FAT push")
    require(re.search(r"\bfat\.(?:push|extend|extend_from_slice)\b", after_body) is not None,
            "candidate FAT loop has no explicit insertion into the reserved vector")


def verify_verifier_sources() -> None:
    path = HERE / "verifier-sources.json"
    require(path.is_file(), "verifier-sources.json is required for replay custody")
    entries = load(path)
    require(isinstance(entries, dict) and entries, "verifier-sources.json must be a non-empty object")
    for relative, digest in entries.items():
        check_hash(digest, f"verifier source {relative}")
        require(isinstance(relative, str), "verifier source path must be a string")
        require(sha(repository_file(relative)) == digest, relative)


def verify_adr_manifest() -> None:
    manifest = load(HERE / "adr-manifest.json")
    require(isinstance(manifest, dict), "ADR manifest must be an object")
    files = manifest.get("files")
    require(isinstance(files, dict) and len(files) == 30,
            "ADR manifest must contain the 30 reviewed files")
    for relative, digest in files.items():
        check_hash(digest, f"ADR {relative}")
        require(sha(repository_file(relative)) == digest, f"ADR hash changed: {relative}")


def verify_sources() -> tuple[dict[str, str], dict[str, str], list[str]]:
    before = source_manifest("before")
    after = source_manifest("after")
    previous = HERE.parent / "change-0510" / "source-manifest.json"
    require(previous.is_file(), "0510 source manifest is required as the preserved base")
    require(before == load(previous), "0511 before source manifest does not preserve 0510 final")

    plan = load(HERE / "plan.json")
    require(isinstance(plan, dict), "plan.json must be an object")
    base_revision = plan.get("base_revision")
    require(base_revision == BASE_REVISION,
            f"plan base revision {base_revision!r} differs from {BASE_REVISION}")
    require(git_blob_sha(base_revision, FAT_SOURCE) == before.get(FAT_SOURCE),
            "before FAT source hash does not match the declared base revision")

    changed = sorted(
        path for path in set(before) | set(after) if before.get(path) != after.get(path)
    )
    require(changed == [FAT_SOURCE], f"source change set is not the single FAT file: {changed!r}")
    for relative, digest in after.items():
        require(sha(repository_file(relative)) == digest,
                f"after source hash changed in checkout: {relative}")
    verify_fat_source_change(base_revision)
    verify_verifier_sources()
    verify_adr_manifest()
    return before, after, changed


def verify_build(stage: str, allocator: bool = False) -> dict[str, Any]:
    directory = stage_dir(stage) if allocator else (HERE / "before" if stage == "before" else HERE)
    receipt = load(build_receipt_path(stage, allocator))
    require(isinstance(receipt, dict), f"{stage} build receipt must be an object")
    label = f"{stage} {'allocator ' if allocator else ''}build"
    require(receipt.get("exit_code") == 0, f"{label} did not exit successfully")
    require(receipt.get("source_unchanged") is True, f"{label} source was not unchanged")
    check_hash(receipt.get("binary_sha256"), f"{label}.binary_sha256")
    require(receipt.get("source_manifest_sha256") == sha(manifest_path(stage)),
            f"{label} source manifest hash is not bound")
    log_name = "allocator-build.log" if allocator else "build.log"
    require(receipt.get("log_sha256") == sha(directory / log_name),
            f"{label} log hash is not bound")
    command = receipt.get("command")
    require(isinstance(command, list) and command and all(isinstance(item, str) for item in command),
            f"{label}.command is not an argv list")
    return receipt


def finite_number(value: Any, context: str) -> float:
    require(isinstance(value, (int, float)) and not isinstance(value, bool),
            f"{context} must be numeric")
    number = float(value)
    require(math.isfinite(number), f"{context} must be finite")
    return number


def verify_metric_vector(value: Any, samples: int, context: str) -> None:
    require(isinstance(value, dict), f"{context} must be a metric vector")
    status = value.get("status")
    require(status in {"measured", "not_applicable", "unavailable", "overflow"},
            f"{context}.status is invalid")
    require(isinstance(value.get("scope"), str) and value["scope"],
            f"{context}.scope is missing")
    values = value.get("values")
    if status == "measured":
        require(isinstance(values, list) and len(values) == samples,
                f"{context}.values must contain {samples} measured values")
        for index, item in enumerate(values):
            if isinstance(item, str):
                require(item in {"sequential", "random", "unknown"},
                        f"{context}.values[{index}] is not a known read pattern")
            else:
                require(isinstance(item, int) and not isinstance(item, bool) and item >= 0,
                        f"{context}.values[{index}] must be a non-negative integer")
    else:
        require(values is None, f"{context} has values despite status {status}")


def verify_metric_vectors(value: Any, samples: int, context: str) -> None:
    """Validate every nested operation metric vector, including allocator data."""

    if isinstance(value, dict):
        if "status" in value and "scope" in value and (
            "values" in value or set(value) <= {"status", "scope", "values"}
        ):
            verify_metric_vector(value, samples, context)
            return
        for key, child in value.items():
            verify_metric_vectors(child, samples, f"{context}.{key}")
    elif isinstance(value, list):
        for index, child in enumerate(value):
            verify_metric_vectors(child, samples, f"{context}[{index}]")


def verify_operation_metrics(row: dict[str, Any], samples: int, context: str,
                             allocator: bool, sample_order: list[int] | None = None) -> None:
    operation = row.get("operation_metrics")
    require(isinstance(operation, dict), f"{context}.operation_metrics is missing")
    require(operation.get("sample_count") == samples,
            f"{context}.operation_metrics.sample_count is not {samples}")
    indices = operation.get("sample_indices")
    require(isinstance(indices, list) and sorted(indices) == list(range(samples)),
            f"{context}.operation_metrics.sample_indices is not a complete permutation")
    if sample_order is not None:
        require(indices == sample_order,
                f"{context}.operation_metrics.sample_indices is not aligned to elapsed_ns")
    require(operation.get("alignment") == "elapsed_ns.samples_by_elapsed_then_sample_index",
            f"{context}.operation_metrics alignment is not explicit")
    verify_metric_vectors(operation, samples, f"{context}.operation_metrics")

    allocation = operation.get("allocation")
    if allocator:
        require(isinstance(allocation, dict), f"{context} allocator metrics are missing")
        require(allocation.get("status") == "measured",
                f"{context} allocator metrics are not measured")
        require(allocation.get("scope") == "operation_global_system_allocator",
                f"{context} allocator scope is not operation-global-system-allocator")
        for field in ALLOCATOR_VECTOR_FIELDS:
            vector = allocation.get(field)
            require(vector is not None, f"{context} allocator field {field} is missing")
            require(vector.get("status") == "measured",
                    f"{context} allocator field {field} is not measured")
    else:
        # Normal XLS rows bracket an explicit unavailable allocator sample;
        # CFB rows have no operation-metrics envelope at all and are checked
        # by the caller.  Never turn either state into a zero allocation.
        if allocation is not None:
            require(allocation.get("status") == "unavailable",
                    f"{context} normal allocator metrics must be unavailable")


def verify_report_identity(report: dict[str, Any], build: dict[str, Any],
                           samples: int, warmup: int, allocator: bool,
                           context: str) -> None:
    require(report.get("schema_version") == 1, f"{context} schema version is not 1")
    tool = report.get("tool")
    require(isinstance(tool, dict), f"{context}.tool is missing")
    expected_instrumentation = (
        "system_allocator_operation_scoped" if allocator else "none"
    )
    require(tool.get("instrumentation") == expected_instrumentation,
            f"{context} instrumentation identity is wrong")
    if allocator:
        require(tool.get("binary") == "litchi-perf-baseline-alloc",
                f"{context} allocator binary identity is wrong")
        require(tool.get("allocator_counter_revision") == "serialized_region_peak_v3",
                f"{context} allocator counter revision is not bound")
    else:
        require(tool.get("binary") == "litchi-perf-baseline",
                f"{context} normal binary identity is wrong")
        require(tool.get("allocator_counter_revision") is None,
                f"{context} normal report carries allocator counter revision")
    require(tool.get("profile") == "release", f"{context} is not a release profile")

    identity = report.get("binary_identity")
    require(isinstance(identity, dict), f"{context}.binary_identity is missing")
    require(identity.get("binary_sha256") == build.get("binary_sha256"),
            f"{context} report binary hash does not match its build")
    check_hash(identity.get("binary_sha256"), f"{context}.binary_identity.binary_sha256")
    require(identity.get("profile") == "release", f"{context} binary profile is not release")
    require(identity.get("executable") is True, f"{context} binary is not executable")

    configuration = report.get("configuration")
    require(isinstance(configuration, dict), f"{context}.configuration is missing")
    require(configuration.get("samples_per_case") == samples,
            f"{context} sample count is not {samples}")
    require(configuration.get("warmup_iterations_per_case") == warmup,
            f"{context} warmup count is not {warmup}")
    require(configuration.get("filesystem_process_isolated") is True,
            f"{context} is not process isolated")
    require(configuration.get("filesystem_fresh_child_per_sample") is True,
            f"{context} does not use a fresh child per sample")
    environment = report.get("environment")
    require(isinstance(environment, dict), f"{context}.environment is missing")
    expected_allocator = (
        "CountingSystemAllocator(std::alloc::System)" if allocator else "Rust system allocator"
    )
    require(environment.get("allocator") == expected_allocator,
            f"{context} allocator environment identity is wrong")


def verify_report_case_configuration(report: dict[str, Any], kind: str, context: str) -> None:
    configuration = report["configuration"]
    cases = configuration.get("cases")
    if kind == "xls":
        require(cases == list(XLS_CASES), f"{context} cases do not match the nine XLS selectors")
    elif kind == "guard":
        require(cases == list(CFB_CASES), f"{context} cases do not match cfb_open")
        require(configuration.get("corpus_shapes") == ["tiny", "few-large"],
                f"{context} guard shapes are not tiny,few-large")
        require(configuration.get("payload_kinds") == ["incompressible"],
                f"{context} guard payload is not incompressible")
    else:
        fail(f"unknown report kind {kind!r}")


def verify_xls_source_evidence(row: dict[str, Any], samples: int, context: str) -> None:
    source = row.get("source")
    if row["case"] not in SOURCE_XLS_CASES:
        require(source is None, f"{context} unexpectedly publishes source-backed evidence")
        return
    require(isinstance(source, dict), f"{context} source evidence is missing")
    one_cell = row["case"].endswith("one_cell")
    expected_counters = {
        "read_calls": 362 if one_cell else 334,
        "read_bytes": 138593 if one_cell else 138459,
        "ordinary_payload_read_calls": 0,
        "ordinary_payload_read_bytes": 0,
        "max_in_flight_reads": 1,
    }
    for field, expected in expected_counters.items():
        require(source.get(field) == [expected] * samples,
                f"{context}.source.{field} changed from the matched logical-read contract")
    xls = source.get("xls")
    require(isinstance(xls, dict), f"{context}.source.xls is missing")
    require(xls.get("implementation") == "source-backed",
            f"{context} source implementation is not source-backed")
    expected_operation = {
        "xls_source_backed_open": "open",
        "xls_source_backed_open_list_worksheets": "open+list",
        "xls_source_backed_open_one_cell": "open+one-cell",
    }[row["case"]]
    require(xls.get("operation") == expected_operation,
            f"{context} source operation identity is wrong")
    require(xls.get("timing_scope") == expected_operation,
            f"{context} source timing scope is wrong")
    require(xls.get("archive_sha256") == row["corpus"]["archive_sha256"],
            f"{context} source archive hash is not bound to the corpus")
    check_hash(xls.get("archive_sha256"), f"{context}.source.xls.archive_sha256")
    check_hash(xls.get("workbook_stream_sha256"),
               f"{context}.source.xls.workbook_stream_sha256")
    vector_fields = (
        "source_retained_bytes",
        "complete_archive_materialized_bytes",
        "parsed_sheet_counts",
        "parsed_cell_counts",
        "source_version_checks",
        "source_version_stability_verified",
        "cfb_structural_read_calls",
        "cfb_structural_read_bytes",
        "workbook_global_read_calls",
        "workbook_global_read_bytes",
        "selected_worksheet_read_calls",
        "selected_worksheet_read_bytes",
        "unselected_worksheet_read_calls",
        "unselected_worksheet_read_bytes",
        "opaque_payload_read_calls",
        "opaque_payload_read_bytes",
        "open_reads_zero_worksheet_payload",
        "selected_query_reads_only_selected_worksheet",
    )
    for field in vector_fields:
        values = xls.get(field)
        require(isinstance(values, list) and len(values) == samples,
                f"{context}.source.xls.{field} must contain {samples} values")
    require(all(item is True for item in xls["source_version_stability_verified"]),
            f"{context} source version stability gate failed")
    if row["case"].endswith("one_cell"):
        require(all(item is True for item in xls["selected_query_reads_only_selected_worksheet"]),
                f"{context} selected-cell locality gate failed")


def verify_xls_row(row: dict[str, Any], samples: int, context: str,
                  allocator: bool) -> None:
    require(row.get("case") in XLS_CASE_SET, f"{context} has an unexpected XLS case")
    require(row.get("sink") is None, f"{context} unexpectedly has sink evidence")
    output = row.get("output_sha256")
    check_hash(output, f"{context}.output_sha256")
    elapsed = _validate_elapsed(row, samples, context)
    verify_operation_metrics(row, samples, context, allocator, elapsed["sample_order"])
    verify_xls_source_evidence(row, samples, context)


def verify_cfb_row(row: dict[str, Any], samples: int, context: str) -> None:
    require(row.get("case") == "cfb_open", f"{context} is not cfb_open")
    require(row.get("sink") is None, f"{context} CFB open has a fabricated sink")
    require(row.get("source") is None, f"{context} CFB open has a fabricated source")
    require(row.get("output_sha256") is None, f"{context} CFB open has a fabricated output hash")
    require(row.get("operation_metrics") is None,
            f"{context} CFB guard operation metrics must remain unavailable")
    _validate_elapsed(row, samples, context)


def verify_artifact_receipt(stage: str, name: str, build: dict[str, Any],
                            samples: int, warmup: int, kind: str,
                            allocator: bool = False) -> tuple[dict[str, Any], dict[str, Any]]:
    directory = stage_dir(stage)
    receipt = load(directory / f"{name}-receipt.json")
    context = f"{stage}/{name}"
    require(isinstance(receipt, dict), f"{context} receipt must be an object")
    require(receipt.get("exit_code") == 0, f"{context} did not exit successfully")
    require(receipt.get("source_unchanged") is True,
            f"{context} source was not unchanged during capture")
    require(receipt.get("binary_sha256") == build.get("binary_sha256"),
            f"{context} receipt binary hash does not match build")
    require(receipt.get("source_manifest_sha256") == build.get("source_manifest_sha256"),
            f"{context} receipt source hash does not match build")
    check_hash(receipt.get("binary_sha256"), f"{context}.binary_sha256")
    artifacts = receipt.get("artifacts")
    require(isinstance(artifacts, dict) and artifacts,
            f"{context} receipt has no artifact hash map")
    for relative, digest in artifacts.items():
        require(isinstance(relative, str) and relative and not Path(relative).is_absolute(),
                f"{context} artifact path is invalid")
        check_hash(digest, f"{context} artifact {relative}")
        artifact = (directory / relative).resolve()
        require(artifact.parent == directory.resolve(),
                f"{context} artifact escapes stage directory")
        require(sha(artifact) == digest, f"{context} artifact hash changed: {relative}")

    command = receipt.get("command")
    require(isinstance(command, list) and all(isinstance(item, str) for item in command),
            f"{context}.command is not an argv list")
    if kind == "callgrind":
        require(command[:2] == ["valgrind", "--tool=callgrind"],
                f"{context} is not a callgrind invocation")
        require("--collect-atstart=no" in command,
                f"{context} does not disable callgrind collection outside the toggle")
        toggle = "--toggle-collect=*SourceBackedWorkbook::from_read_at_with_limits"
        require(toggle in command, f"{context} does not bind the intended callgrind scope")
        require(receipt.get("scope", "").startswith("SourceBackedWorkbook::from_read_at_with_limits"),
                f"{context} receipt scope is not the constructor-only scope")
        require("excludes selected-cell query and setup" in receipt.get("scope", ""),
                f"{context} receipt does not document excluded setup/query")
    else:
        require(command[:3] == ["taskset", "-c", "2"],
                f"{context} is not pinned to CPU 2")
        expected_scope = (
            "instrumented operation allocation deltas; timings excluded from native comparisons"
            if allocator
            else "existing matched operation clocks; whole-child RSS includes setup, copies and oracles"
        )
        require(receipt.get("scope") == expected_scope,
                f"{context} timing/allocation scope text is not exact")

    report_name = f"{name}-report.json"
    catalog_name = f"{name}-catalog.json"
    require(report_name in artifacts and catalog_name in artifacts,
            f"{context} receipt does not bind report and catalog")
    report = load(directory / report_name)
    catalog = load(directory / catalog_name)
    require(isinstance(report, dict), f"{context} report must be an object")
    require(isinstance(catalog, dict), f"{context} catalog must be an object")
    validate_binding(report, catalog)
    verify_report_identity(report, build, samples, warmup, allocator, context)
    if kind == "xls" or kind == "callgrind":
        expected_count = 1 if kind == "callgrind" else len(XLS_CASES)
        rows = report.get("results")
        require(isinstance(rows, list) and len(rows) == expected_count,
                f"{context} has {len(rows) if isinstance(rows, list) else 'no'} rows; expected {expected_count}")
        for index, row in enumerate(rows):
            require(isinstance(row, dict), f"{context} row {index} is not an object")
            if kind == "callgrind":
                require(row.get("case") == CALLGRIND_CASE,
                        f"{context} callgrind selector is not {CALLGRIND_CASE}")
                elapsed = _validate_elapsed(row, samples, f"{context}.row[{index}]")
                verify_operation_metrics(row, samples, f"{context}.row[{index}]", allocator,
                                         elapsed["sample_order"])
                verify_xls_source_evidence(row, samples, f"{context}.row[{index}]")
                check_hash(row.get("output_sha256"), f"{context}.output_sha256")
            else:
                verify_xls_row(row, samples, f"{context}.row[{index}]", allocator)
    elif kind == "guard":
        rows = report.get("results")
        require(isinstance(rows, list) and len(rows) == 2,
                f"{context} must contain exactly two CFB guard rows")
        for index, row in enumerate(rows):
            require(isinstance(row, dict), f"{context} row {index} is not an object")
            verify_cfb_row(row, samples, f"{context}.row[{index}]")
    else:
        fail(f"unknown capture kind {kind!r}")
    return receipt, report


def row_key(row: dict[str, Any]) -> tuple[str, str]:
    return row["case"], canonical(row["corpus"])


def identity_fields(row: dict[str, Any]) -> dict[str, Any]:
    source = row.get("source")
    source_identity = None
    if isinstance(source, dict) and isinstance(source.get("xls"), dict):
        xls = source["xls"]
        source_identity = {
            key: xls.get(key)
            for key in (
                "implementation",
                "operation",
                "timing_scope",
                "archive_sha256",
                "workbook_stream_sha256",
            )
        }
    return {
        "corpus": row.get("corpus"),
        "sink": row.get("sink"),
        "output_sha256": row.get("output_sha256"),
        "source": source_identity,
    }


def verify_identity_matrix(reports: dict[tuple[str, str, str], dict[str, Any]],
                           allocator_reports: dict[tuple[str, str], dict[str, Any]],
                           callgrind_reports: dict[tuple[str, str], dict[str, Any]]) -> None:
    native_rows: dict[tuple[str, str, str], dict[str, Any]] = {}
    for (lane, stage, repeat), report in reports.items():
        rows = report["results"]
        for row in rows:
            key = row_key(row)
            require((lane, stage, repeat, key) not in native_rows,
                    f"duplicate native identity {lane}/{stage}/{repeat}/{key[0]}")
            native_rows[lane, stage, repeat, key] = row

    for lane, expected_cases in (("xls", XLS_CASES), ("guard", CFB_CASES)):
        keys = {
            key
            for (report_lane, _stage, _repeat, key) in native_rows
            if report_lane == lane
        }
        require(len(keys) == (len(expected_cases) if lane == "xls" else 2),
                f"{lane} identity matrix does not have the expected corpus keys")
        for key in keys:
            reference = native_rows[lane, "before", "r1", key]
            for stage in ("before", "after"):
                for repeat in ("r1", "r2"):
                    current = native_rows[lane, stage, repeat, key]
                    require(identity_fields(current) == identity_fields(reference),
                            f"{lane} identity changed at {stage}/{repeat}/{key[0]}")

    # Allocator reports use the same nine XLS corpora and output oracles.  The
    # allocation target's elapsed vectors remain a separate evidence stream.
    for (stage, repeat), report in allocator_reports.items():
        for row in report["results"]:
            key = row_key(row)
            native = native_rows["xls", stage, repeat, key]
            require(identity_fields(row) == identity_fields(native),
                    f"allocator/native identity differs at {stage}/{repeat}/{key[0]}")

    # Callgrind has one row per stage and uses the same owned-source one-cell
    # oracle as its native counterpart.  Its instruction count is diagnostic,
    # not an elapsed-time comparison.
    for stage, report in callgrind_reports.items():
        row = report["results"][0]
        key = row_key(row)
        native = native_rows["xls", stage, "r1", key]
        require(identity_fields(row) == identity_fields(native),
                f"callgrind/native identity differs at {stage}/{key[0]}")


def verify_callgrind_raw(stage: str) -> int:
    path = stage_dir(stage) / "callgrind.out"
    text = path.read_text(encoding="utf-8", errors="strict")
    match = CALLGRIND_SUMMARY_RE.search(text)
    require(match is not None, f"{stage}/callgrind.out has no summary")
    value = int(match.group(1))
    require(value > 0, f"{stage}/callgrind.out has an empty summary")
    return value


def rss(path: Path) -> int:
    text = path.read_text(encoding="utf-8", errors="strict")
    match = RSS_RE.search(text)
    require(match is not None, f"{path} has no /usr/bin/time RSS line")
    value = int(match.group(1))
    require(value > 0, f"{path} reports non-positive RSS")
    return value


def percent(after: float, before: float) -> float:
    require(before > 0, "cannot compute a delta from a non-positive baseline")
    return (after / before - 1.0) * 100.0


def paired_delta(before: dict[str, Any], after: dict[str, Any],
                 repeat: str, case: str, corpus: dict[str, Any]) -> dict[str, Any]:
    b = before["elapsed_ns"]
    a = after["elapsed_ns"]
    values: dict[str, Any] = {
        "repeat": repeat,
        "case": case,
        "corpus": corpus,
        "before_p50_ns": b["p50"],
        "after_p50_ns": a["p50"],
    }
    for metric in ELAPSED_METRICS:
        values[f"{metric}_change_percent"] = percent(float(a[metric]), float(b[metric]))
    values["throughput_change_percent"] = percent(float(b["mean"]), float(a["mean"]))
    values["adverse_latency_or_throughput_over_5_percent"] = (
        any(values[f"{metric}_change_percent"] > 5.0 for metric in ELAPSED_METRICS)
        or values["throughput_change_percent"] < -5.0
    )
    return values


def verify_native_comparisons(reports: dict[tuple[str, str, str], dict[str, Any]]) -> tuple[list[dict[str, Any]], list[dict[str, Any]], list[dict[str, Any]]]:
    pairs: list[dict[str, Any]] = []
    drift: list[dict[str, Any]] = []
    memory: list[dict[str, Any]] = []

    for lane in ("xls", "guard"):
        row_maps: dict[tuple[str, str], dict[tuple[str, str], dict[str, Any]]] = {}
        for stage in ("before", "after"):
            for repeat in ("r1", "r2"):
                report = reports[lane, stage, repeat]
                row_maps[stage, repeat] = {row_key(row): row for row in report["results"]}
        for repeat in ("r1", "r2"):
            for key, before in row_maps["before", repeat].items():
                after = row_maps["after", repeat].get(key)
                require(after is not None, f"{lane} after row missing for {key[0]}")
                pairs.append(paired_delta(before, after, repeat, key[0], before["corpus"]))
        for stage in ("before", "after"):
            for key, first in row_maps[stage, "r1"].items():
                second = row_maps[stage, "r2"].get(key)
                require(second is not None, f"{lane} repeat row missing for {key[0]}")
                changes = {
                    metric: percent(float(second["elapsed_ns"][metric]),
                                    float(first["elapsed_ns"][metric]))
                    for metric in ELAPSED_METRICS
                }
                drift.append({
                    "lane": lane,
                    "stage": stage,
                    "case": key[0],
                    "corpus": first["corpus"],
                    "repeat_change_percent": changes,
                    "exceeds_drift_ceiling": any(
                        abs(changes[metric]) > limit for metric, limit in DRIFT_LIMITS.items()
                    ),
                })
        log_name = f"{lane}-{{repeat}}.log"
        for repeat in ("r1", "r2"):
            b = rss(stage_dir("before") / log_name.format(repeat=repeat))
            a = rss(stage_dir("after") / log_name.format(repeat=repeat))
            change = percent(float(a), float(b))
            memory.append({
                "lane": lane,
                "repeat": repeat,
                "scope": "whole child: selected cases, setup, copies, and post-operation oracles",
                "before_peak_rss_kib": b,
                "after_peak_rss_kib": a,
                "change_percent": change,
                "adverse_over_5_percent": change > 5.0,
            })
    return pairs, drift, memory


def verify_negative_probe(row: dict[str, Any], samples: int) -> None:
    bad = copy.deepcopy(row)
    bad["elapsed_ns"]["samples"].pop()
    try:
        _validate_elapsed(bad, samples, "negative short vector")
    except (ValueError, AssertionError):
        return
    fail("negative short elapsed vector was accepted")


def verify_allocator_reports(builds: dict[str, dict[str, Any]],
                             native_reports: dict[tuple[str, str, str], dict[str, Any]]) -> dict[tuple[str, str], dict[str, Any]]:
    allocator_reports: dict[tuple[str, str], dict[str, Any]] = {}
    for stage, repeat in (("before", "r1"), ("after", "r1"), ("after", "r2"), ("before", "r2")):
        receipt, report = verify_artifact_receipt(
            stage,
            f"allocator-{repeat}",
            builds[stage, "allocator"],
            samples=30,
            warmup=3,
            kind="xls",
            allocator=True,
        )
        del receipt
        allocator_reports[stage, repeat] = report
        native = native_reports["xls", stage, repeat]
        require(
            {row_key(row) for row in report["results"]}
            == {row_key(row) for row in native["results"]},
            f"allocator {stage}/{repeat} rows do not cover native XLS identities",
        )
    return allocator_reports


def verify_capture_order(receipts: dict[tuple[str, str, str], dict[str, Any]],
                         allocator_receipts: dict[tuple[str, str], dict[str, Any]]) -> None:
    expected = (("before", "r1"), ("after", "r1"), ("after", "r2"), ("before", "r2"))
    for lane in ("xls", "guard"):
        starts = []
        for stage, repeat in expected:
            value = receipts[lane, stage, repeat].get("started_utc")
            require(isinstance(value, str) and value, f"{lane}/{stage}/{repeat} has no start time")
            try:
                starts.append(_datetime.datetime.fromisoformat(value.replace("Z", "+00:00")))
            except ValueError as error:
                fail(f"{lane}/{stage}/{repeat} has invalid start time: {error}")
        require(starts == sorted(starts) and len(set(starts)) == 4,
                f"{lane} captures are not serial ABBA")
    starts = []
    for stage, repeat in expected:
        value = allocator_receipts[stage, repeat].get("started_utc")
        require(isinstance(value, str) and value, f"allocator/{stage}/{repeat} has no start time")
        try:
            starts.append(_datetime.datetime.fromisoformat(value.replace("Z", "+00:00")))
        except ValueError as error:
            fail(f"allocator/{stage}/{repeat} has invalid start time: {error}")
    require(starts == sorted(starts) and len(set(starts)) == 4,
            "allocator captures are not serial ABBA")


def main() -> None:
    before_manifest, after_manifest, changed = verify_sources()
    del before_manifest, after_manifest

    builds: dict[tuple[str, str], dict[str, Any]] = {}
    for stage in ("before", "after"):
        builds[stage, "normal"] = verify_build(stage)
        builds[stage, "allocator"] = verify_build(stage, allocator=True)

    reports: dict[tuple[str, str, str], dict[str, Any]] = {}
    receipts: dict[tuple[str, str, str], dict[str, Any]] = {}
    for lane, samples, warmup in (("xls", 1000, 20), ("guard", 1000, 20)):
        for stage, repeat in (("before", "r1"), ("after", "r1"), ("after", "r2"), ("before", "r2")):
            receipt, report = verify_artifact_receipt(
                stage,
                f"{lane}-{repeat}",
                builds[stage, "normal"],
                samples=samples,
                warmup=warmup,
                kind=lane,
            )
            verify_report_case_configuration(report, lane, f"{stage}/{lane}-{repeat}")
            reports[lane, stage, repeat] = report
            receipts[lane, stage, repeat] = receipt

    callgrind_reports: dict[str, dict[str, Any]] = {}
    callgrind_receipts: dict[str, dict[str, Any]] = {}
    callgrind_summaries: dict[str, int] = {}
    for stage in ("before", "after"):
        receipt, report = verify_artifact_receipt(
            stage,
            "callgrind",
            builds[stage, "normal"],
            samples=5,
            warmup=0,
            kind="callgrind",
        )
        callgrind_receipts[stage] = receipt
        callgrind_reports[stage] = report
        callgrind_summaries[stage] = verify_callgrind_raw(stage)

    allocator_reports = verify_allocator_reports(builds, reports)
    allocator_receipts = {
        (stage, repeat): load(stage_dir(stage) / f"allocator-{repeat}-receipt.json")
        for stage, repeat in (("before", "r1"), ("after", "r1"), ("after", "r2"), ("before", "r2"))
    }
    verify_capture_order(receipts, allocator_receipts)
    verify_identity_matrix(reports, allocator_reports, callgrind_reports)
    pairs, drift, memory = verify_native_comparisons(reports)
    verify_negative_probe(reports["xls", "after", "r1"]["results"][0], 1000)

    allocation_identity = {
        f"{stage}/{repeat}": {
            "rows": len(report["results"]),
            "samples_per_row": 30,
            "warmup_iterations_per_case": 3,
            "elapsed_comparison": "excluded from native deltas",
            "allocation_vectors": list(ALLOCATOR_VECTOR_FIELDS),
        }
        for (stage, repeat), report in allocator_reports.items()
    }
    output = {
        "source_changes": changed,
        "native": {
            "xls_rows_per_report": 9,
            "cfb_rows_per_report": 2,
            "samples_per_row": 1000,
            "warmup_iterations_per_case": 20,
            "xls_total_elapsed_samples": 9 * 4 * 1000,
            "cfb_total_elapsed_samples": 2 * 4 * 1000,
            "abba_order": ["before/r1", "after/r1", "after/r2", "before/r2"],
        },
        "allocator": {
            "rows_per_report": 9,
            "samples_per_row": 30,
            "warmup_iterations_per_case": 3,
            "total_elapsed_samples": 9 * 4 * 30,
            "instrumented_elapsed_excluded_from_native_comparison": True,
            "reports": allocation_identity,
        },
        "callgrind": {
            "case": CALLGRIND_CASE,
            "samples_per_report": 5,
            "warmup_iterations_per_case": 0,
            "collection_scope": "SourceBackedWorkbook::from_read_at_with_limits",
            "instruction_reference_summaries": callgrind_summaries,
            "excluded_from_native_comparison": True,
        },
        "pairs": pairs,
        "drift": drift,
        "whole_child_rss": memory,
        "adverse_flags_are_observations": True,
        "negative_short_vector_rejected": True,
        "corpus_and_output_identities_verified": True,
        "report_catalog_bindings_verified": True,
        "receipt_binary_and_source_hashes_verified": True,
        "capture_boundaries_verified": True,
    }
    from verify_allocation import validate as validate_allocation
    from verify_hardware import validate as validate_hardware
    from verify_supplemental import validate as validate_supplemental
    output['allocation'] = validate_allocation(HERE)
    output['hardware'] = validate_hardware(HERE)
    output['supplemental'] = validate_supplemental(HERE)
    print(json.dumps(output, indent=2, sort_keys=True))


if __name__ == "__main__":
    try:
        main()
    except VerificationError as error:
        print(f"verify.py: {error}", file=sys.stderr)
        raise SystemExit(1)
