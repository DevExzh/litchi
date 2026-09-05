#!/usr/bin/env python3
"""Bounded, read-only verifier for the 0410 XLSX ABBA capture.

The capture itself is produced by run-leg.py.  This verifier only reads the
four selected-cell normal reports, four edit/save guardrail reports, four
selected-cell allocator reports, their catalog sidecars, command journal,
environment files, and build identity files.  It validates provenance and
operation-scoped allocation vectors, then writes a compact JSON observation.
It deliberately does not compare allocator elapsed time.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import statistics
import subprocess
import sys
from pathlib import Path
from typing import Any

SCRIPT_VERSION = "0410-allocation-result-verifier-v1"
DEFAULT_OUT = Path("/tmp/litchi-goal-0410-capture")
DEFAULT_BINARY_ROOT = Path("/tmp/litchi-goal-0410-binaries")

LEGS = ("a1", "b1", "b2", "a2")
CONTROL_LEGS = ("a1", "a2")
CANDIDATE_LEGS = ("b1", "b2")
ALL_KINDS = ("normal", "guards", "allocator")
SELECTED_CASE = "xlsx_file_selected_cell"
GUARD_CASES = (
    "xlsx_source_backed_cell_values_one_edit_save",
    "xlsx_eager_cell_values_one_edit_save",
)
ALLOCATOR_METRICS = (
    "allocation_calls",
    "deallocation_calls",
    "reallocation_calls",
    "failed_allocation_calls",
    "allocated_bytes",
    "deallocated_bytes",
)
ALLOCATOR_SCOPE = "operation_global_system_allocator"
RUSTFLAGS = "-C force-frame-pointers=yes -C force-unwind-tables=yes"
RUSTC_PREFIX = "rustc 1.98.1 "
CONTROL_REVISION = "972dc25be0dbd6690c74429839a48288d637e2d5"
CANDIDATE_REVISION = "e4d477466718a8fad38cd55b9babe0b826e7f3a7"
ARCHIVE_SHA256 = "dfff7ec0c749d9e404091776f15a8fb690985af7f58efdfe659dbeaed7145036"
INPUT_SEMANTIC_SHA256 = (
    "020fdd140d2959ea4f480676a3d4d0bf840927e25251cb6cad37a043ab80627e"
)
OUTPUT_SHA256 = "9b7b66a02007eeb63498fd5de4c6b7115ace0383ce37d97e1a9560ef7bfadec1"
OUTPUT_SEMANTIC_SHA256 = (
    "3cd21160d4f74fa0f097ab40be08e211b3e460cea788aa2b6705a55fdece07de"
)
UNTOUCHED_MEMBER_SHA256 = (
    "7105fcbce160328f666e69fcfd18da9e19fd71dd7b63961e7cddd29d5da1a17d"
)
SELECTED_CELL = {
    "canonical_sheet_name": "Bench01",
    "sheet_position": 1,
    "prepared_selector": "bEnCh01",
    "cell_address": "M29",
    "view": "stored",
    "value_kind": "number",
    "lexical_value": "1028012",
    "digest": "36e53d9002ae8c433ad918b400196fb886fa675f850076808ac51327d1f42ac1",
}
CORPUS = {
    "name": "xlsx-cell-values-medium",
    "generator": "litchi-xlsx-cell-values-source-edit-media-multi-sheet-v1",
    "package_format": "XLSX/OPC/ZIP",
    "shape": "medium",
    "payload_kind": "deterministic-multi-sheet-scalar-grid-with-media",
    "compression": "deflate",
    "entry_count": 9216,
    "archive_member_count": 17,
    "entry_bytes": 4,
    "uncompressed_payload_bytes": 4231168,
    "archive_bytes": 4226429,
    "archive_sha256": ARCHIVE_SHA256,
    "target_entry": "Sheet1!A1",
    "target_payload_bytes": 1,
    "target_payload_sha256": "5feceb66ffc86f38d952786c6d696c79c2dbc239dd4e91b46729d73a27fb57e9",
    "xlsx": {
        "sheet_count": 4,
        "rows_per_sheet": 48,
        "columns_per_sheet": 48,
        "one_percent_update_count": 93,
        "source_members": {
            "shared_strings": None,
            "styles": "xl/styles.xml",
            "workbook": "xl/workbook.xml",
            "worksheets": [
                "xl/worksheets/sheet1.xml",
                "xl/worksheets/sheet2.xml",
                "xl/worksheets/sheet3.xml",
                "xl/worksheets/sheet4.xml",
            ],
        },
    },
}


class VerificationError(ValueError):
    """The capture is incomplete, malformed, or fails an identity gate."""


def reject_constant(token: str) -> Any:
    raise VerificationError(f"non-finite JSON constant {token!r}")


def no_duplicate_pairs(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for key, value in pairs:
        if key in result:
            raise VerificationError(f"duplicate JSON key {key!r}")
        result[key] = value
    return result


def canonical_bytes(value: Any) -> bytes:
    return json.dumps(
        value,
        sort_keys=True,
        separators=(",", ":"),
        ensure_ascii=True,
        allow_nan=False,
    ).encode("utf-8")


def sha256_bytes(value: bytes) -> str:
    return hashlib.sha256(value).hexdigest()


def load_json(path: Path) -> tuple[Any, dict[str, Any]]:
    source_path = path
    if not source_path.exists() and Path(str(source_path) + ".zst").exists():
        source_path = Path(str(source_path) + ".zst")
    try:
        artifact = source_path.read_bytes()
    except OSError as error:
        raise VerificationError(f"cannot read {source_path}: {error}") from error
    if source_path.suffix == ".zst":
        try:
            raw = subprocess.check_output(["zstd", "-q", "-dc", str(source_path)])
        except (OSError, subprocess.CalledProcessError) as error:
            raise VerificationError(f"cannot decompress {source_path}: {error}") from error
    else:
        raw = artifact
    try:
        value = json.loads(
            raw.decode("utf-8"),
            object_pairs_hook=no_duplicate_pairs,
            parse_constant=reject_constant,
        )
    except (UnicodeError, json.JSONDecodeError, VerificationError) as error:
        raise VerificationError(f"cannot parse {source_path}: {error}") from error
    return value, {
        "path": str(source_path),
        "bytes": len(raw),
        "sha256": sha256_bytes(raw),
        "canonical_sha256": sha256_bytes(canonical_bytes(value)),
        **(
            {
                "artifact_bytes": len(artifact),
                "artifact_sha256": sha256_bytes(artifact),
                "compressed": True,
            }
            if source_path.suffix == ".zst"
            else {"compressed": False}
        ),
    }


def expect(condition: bool, message: str) -> None:
    if not condition:
        raise VerificationError(message)


def obj(value: Any, context: str) -> dict[str, Any]:
    expect(isinstance(value, dict), f"{context} must be an object")
    return value


def list_value(value: Any, context: str) -> list[Any]:
    expect(isinstance(value, list), f"{context} must be a list")
    return value


def integer(value: Any, context: str, *, minimum: int = 0) -> int:
    expect(isinstance(value, int) and not isinstance(value, bool), f"{context} must be an integer")
    expect(value >= minimum, f"{context} must be >= {minimum}")
    return value


def constant_vector(value: Any, count: int, context: str, expected: Any = None) -> list[Any]:
    values = list_value(value, context)
    expect(len(values) == count, f"{context} has {len(values)} values; expected {count}")
    if values:
        expect(len({canonical_bytes(item) for item in values}) == 1, f"{context} is not constant")
        if expected is not None:
            expect(values[0] == expected, f"{context}[0]={values[0]!r}; expected {expected!r}")
    return values


def report_path(out: Path, leg: str, kind: str) -> Path:
    return out / f"{leg}-{kind}.json"


def catalog_path(out: Path, leg: str, kind: str) -> Path:
    return out / f"{leg}-{kind}.catalog.json"


def expected_cases(kind: str) -> tuple[str, ...]:
    if kind == "normal":
        return (SELECTED_CASE,)
    if kind == "guards":
        return GUARD_CASES
    if kind == "allocator":
        return (SELECTED_CASE,)
    raise VerificationError(f"unknown capture kind {kind!r}")


def expected_samples(kind: str) -> tuple[int, int]:
    return (30, 3) if kind == "allocator" else (500, 20)


def expected_marker(kind: str) -> str | None:
    return "compressed-member-intersections-v1" if kind == "guards" else None


def role_for_leg(leg: str) -> str:
    return "control" if leg in CONTROL_LEGS else "candidate"


def expected_revision(role: str) -> str:
    return CONTROL_REVISION if role == "control" else CANDIDATE_REVISION


def load_role_identity(out: Path, role: str) -> tuple[dict[str, Any], dict[str, Any]]:
    candidates = (
        out / f"{role}-build-identity.json",
        DEFAULT_BINARY_ROOT / role / "identity.json",
    )
    for path in candidates:
        if path.exists():
            value, info = load_json(path)
            identity = obj(value, str(path))
            expect(identity.get("revision") == expected_revision(role), f"{path}: revision mismatch")
            expect(identity.get("status") == "", f"{path}: build worktree is not clean")
            binaries = obj(identity.get("binaries"), f"{path}.binaries")
            for binary_name in ("litchi-perf-baseline", "litchi-perf-baseline-alloc"):
                meta = obj(binaries.get(binary_name), f"{path}.binaries.{binary_name}")
                integer(meta.get("bytes"), f"{path}.binaries.{binary_name}.bytes", minimum=1)
                digest = meta.get("sha256")
                expect(isinstance(digest, str) and len(digest) == 64, f"{path}: invalid {binary_name} hash")
            return identity, info
    raise VerificationError(f"missing {role} build identity (looked in {candidates!r})")


def validate_report_header(
    report: dict[str, Any],
    label: str,
    kind: str,
    identity: dict[str, Any],
) -> dict[str, Any]:
    role = role_for_leg(label.split("-", 1)[0])
    samples, warmups = expected_samples(kind)
    cases = list(expected_cases(kind))
    expect(report.get("schema_version") == 1, f"{label}: schema_version must be 1")
    tool = obj(report.get("tool"), f"{label}.tool")
    expect(tool.get("name") == "litchi-perf-baseline", f"{label}: tool name mismatch")
    expected_binary = "litchi-perf-baseline-alloc" if kind == "allocator" else "litchi-perf-baseline"
    expect(tool.get("binary") == expected_binary, f"{label}: tool binary mismatch")
    expected_instrumentation = "system_allocator_operation_scoped" if kind == "allocator" else "none"
    expect(tool.get("instrumentation") == expected_instrumentation, f"{label}: instrumentation mismatch")

    environment = obj(report.get("environment"), f"{label}.environment")
    expect(environment.get("git_revision") == expected_revision(role), f"{label}: git revision mismatch")
    expect(environment.get("git_worktree_dirty") is False, f"{label}: worktree must be clean")
    expect(environment.get("cpu_affinity") == "2", f"{label}: CPU affinity must be 2")
    expect(environment.get("rustflags") == RUSTFLAGS, f"{label}: RUSTFLAGS mismatch")
    rustc = environment.get("rustc_version")
    expect(isinstance(rustc, str) and rustc.startswith(RUSTC_PREFIX), f"{label}: rustc must be 1.98.1")
    expected_allocator = "CountingSystemAllocator(std::alloc::System)" if kind == "allocator" else "Rust system allocator"
    expect(environment.get("allocator") == expected_allocator, f"{label}: allocator identity mismatch")

    binary_identity = obj(report.get("binary_identity"), f"{label}.binary_identity")
    binary_name = expected_binary
    build_binary = obj(obj(identity.get("binaries"), "identity.binaries").get(binary_name), f"identity.binaries.{binary_name}")
    expect(binary_identity.get("binary_sha256") == build_binary.get("sha256"), f"{label}: binary SHA mismatch")
    expect(binary_identity.get("binary_bytes") == build_binary.get("bytes"), f"{label}: binary byte count mismatch")
    expect(binary_identity.get("profile") == "release", f"{label}: binary profile mismatch")
    expect(binary_identity.get("executable") is True, f"{label}: binary is not executable")

    configuration = obj(report.get("configuration"), f"{label}.configuration")
    expected_configuration = {
        "samples_per_case": samples,
        "warmup_iterations_per_case": warmups,
        "filesystem_cache_states": ["warm"],
        "filesystem_fresh_child_per_sample": True,
        "filesystem_process_isolated": True,
        "filesystem_root_selected": False,
        "cases": cases,
        "xlsx_cell_crud_shapes": ["medium"],
        "execution_workers": [1],
    }
    for key, expected in expected_configuration.items():
        expect(configuration.get(key) == expected, f"{label}: configuration.{key} mismatch")
    marker = expected_marker(kind)
    if marker is None:
        expect("xlsx_cell_values_range_accounting" not in configuration, f"{label}: unexpected range-accounting marker")
    else:
        expect(configuration.get("xlsx_cell_values_range_accounting") == marker, f"{label}: range-accounting marker mismatch")

    results = list_value(report.get("results"), f"{label}.results")
    expect(len(results) == len(cases), f"{label}: result count mismatch")
    result_map: dict[str, dict[str, Any]] = {}
    for index, result_value in enumerate(results):
        result = obj(result_value, f"{label}.results[{index}]")
        case = result.get("case")
        expect(case in cases and case not in result_map, f"{label}: unexpected or duplicate result case {case!r}")
        result_map[case] = result
        expect(result.get("cache_state") == "warm", f"{label}.{case}: cache state mismatch")
        expect(result.get("corpus") == CORPUS, f"{label}.{case}: corpus identity mismatch")
        elapsed = obj(result.get("elapsed_ns"), f"{label}.{case}.elapsed_ns")
        elapsed_samples = list_value(elapsed.get("samples"), f"{label}.{case}.elapsed_ns.samples")
        expect(len(elapsed_samples) == samples, f"{label}.{case}: elapsed sample count mismatch")
        expect(all(isinstance(item, int) and item > 0 for item in elapsed_samples), f"{label}.{case}: elapsed samples must be positive integers")
        order = list_value(elapsed.get("sample_order"), f"{label}.{case}.elapsed_ns.sample_order")
        expect(order == sorted(order) and order == list(range(samples)), f"{label}.{case}: sample order must be 0..N-1")
        operation_metrics = obj(result.get("operation_metrics"), f"{label}.{case}.operation_metrics")
        expect(operation_metrics.get("sample_count") == samples, f"{label}.{case}: operation sample count mismatch")

    return {
        "label": label,
        "kind": kind,
        "role": role,
        "revision": expected_revision(role),
        "binary_sha256": binary_identity["binary_sha256"],
        "raw": None,
        "configuration": configuration,
        "results": result_map,
        "sample_count": samples,
        "warmup_iterations": warmups,
    }


def validate_catalog(
    value: Any,
    label: str,
    cases: tuple[str, ...],
    revision: str,
) -> None:
    catalog = obj(value, label)
    expect(catalog.get("manifest_version") == 2, f"{label}: manifest version mismatch")
    expect(catalog.get("manifest_kind") == "corpus-catalog", f"{label}: manifest kind mismatch")
    expect(catalog.get("catalog_id") == "litchi-perf-corpus-v2", f"{label}: catalog id mismatch")
    build = obj(catalog.get("build"), f"{label}.build")
    expect(build.get("git_revision") == revision, f"{label}: catalog revision mismatch")
    expect(build.get("git_worktree_dirty") is False, f"{label}: catalog worktree must be clean")
    corpora = list_value(catalog.get("corpora"), f"{label}.corpora")
    expect(len(corpora) == 1, f"{label}: expected one corpus")
    corpus_record = obj(corpora[0], f"{label}.corpora[0]")
    bytes_record = obj(corpus_record.get("bytes"), f"{label}.corpora[0].bytes")
    expect(bytes_record.get("archive_sha256") == ARCHIVE_SHA256, f"{label}: catalog archive hash mismatch")
    expect(bytes_record.get("archive_bytes") == CORPUS["archive_bytes"], f"{label}: catalog archive size mismatch")
    bindings = list_value(catalog.get("case_bindings"), f"{label}.case_bindings")
    expect({obj(item, f"{label}.case_bindings[]").get("case") for item in bindings} == set(cases), f"{label}: catalog cases mismatch")
    for item in bindings:
        binding = obj(item, f"{label}.case_bindings[]")
        expect(binding.get("corpus_id") == f"xlsx-opc-zip:sha256:{ARCHIVE_SHA256}", f"{label}: catalog corpus id mismatch")


def validate_selected_filesystem(
    report_info: dict[str, Any],
    label: str,
    *,
    allocator: bool,
    global_child_ids: set[int],
) -> dict[str, Any]:
    report = report_info["_report"]
    samples = report_info["sample_count"]
    filesystem = list_value(report.get("filesystem_evidence"), f"{label}.filesystem_evidence")
    expect(len(filesystem) == 1, f"{label}: expected one selected-cell filesystem envelope")
    envelope = obj(filesystem[0], f"{label}.filesystem_evidence[0]")
    expect(envelope.get("case") == SELECTED_CASE, f"{label}: filesystem case mismatch")
    expect(envelope.get("corpus") == CORPUS, f"{label}: filesystem corpus mismatch")
    expect(envelope.get("warmup_iterations") == report_info["warmup_iterations"], f"{label}: filesystem warmups mismatch")
    expect(envelope.get("sample_count") == samples, f"{label}: filesystem sample count mismatch")
    expect(envelope.get("cache_states") == ["warm"], f"{label}: filesystem cache states mismatch")
    expect(envelope.get("fresh_child_per_sample") is True, f"{label}: fresh-child contract missing")
    raw_samples = list_value(envelope.get("samples"), f"{label}.filesystem_evidence[0].samples")
    expect(len(raw_samples) == samples, f"{label}: filesystem raw sample count mismatch")
    child_ids: set[int] = set()
    allocation_by_index: dict[int, dict[str, int]] = {}
    for ordinal, raw_value in enumerate(raw_samples):
        raw = obj(raw_value, f"{label}.filesystem_evidence.samples[{ordinal}]")
        sample_index = integer(raw.get("sample_index"), f"{label}.filesystem_evidence.samples[{ordinal}].sample_index")
        expect(sample_index == ordinal, f"{label}: sample index/order mismatch at {ordinal}")
        child_id = integer(raw.get("child_process_id"), f"{label}.filesystem_evidence.samples[{ordinal}].child_process_id", minimum=1)
        expect(child_id not in child_ids, f"{label}: duplicate child process id {child_id}")
        expect(child_id not in global_child_ids, f"{label}: child process id reused across ABBA legs: {child_id}")
        child_ids.add(child_id)
        global_child_ids.add(child_id)
        expect(raw.get("cache_state") == "warm", f"{label}: raw cache state mismatch")
        expect(raw.get("xlsx_source_sha256") == ARCHIVE_SHA256, f"{label}: source archive identity mismatch")
        expect(raw.get("xlsx_semantic_sha256") == INPUT_SEMANTIC_SHA256, f"{label}: source semantic identity mismatch")
        expect(raw.get("xlsx_selected_cell") == SELECTED_CELL, f"{label}: selected-cell oracle mismatch")
        if allocator:
            allocation = obj(raw.get("allocation_metrics"), f"{label}.filesystem_evidence.samples[{ordinal}].allocation_metrics")
            expect(allocation.get("status") == "measured", f"{label}: allocation sample is not measured")
            expect(allocation.get("scope") == ALLOCATOR_SCOPE, f"{label}: allocation sample scope mismatch")
            allocation_by_index[sample_index] = {
                metric: integer(allocation.get(metric), f"{label}.allocation.{metric}")
                for metric in ALLOCATOR_METRICS
            }
        else:
            expect("allocation_metrics" not in raw, f"{label}: normal sample unexpectedly has allocation metrics")
    expect(len(child_ids) == samples, f"{label}: child id count mismatch")
    return {"child_process_ids": sorted(child_ids), "allocation_by_index": allocation_by_index}


def validate_selected_result_allocation(
    report_info: dict[str, Any],
    label: str,
    filesystem: dict[str, Any],
) -> dict[str, list[int]]:
    result = report_info["results"][SELECTED_CASE]
    operation = obj(result.get("operation_metrics"), f"{label}.{SELECTED_CASE}.operation_metrics")
    allocation = obj(operation.get("allocation"), f"{label}.{SELECTED_CASE}.operation_metrics.allocation")
    expect(allocation.get("status") == "measured", f"{label}: operation allocation status mismatch")
    expect(allocation.get("scope") == ALLOCATOR_SCOPE, f"{label}: operation allocation scope mismatch")
    samples = report_info["sample_count"]
    vectors: dict[str, list[int]] = {}
    for metric in ALLOCATOR_METRICS:
        metric_value = obj(allocation.get(metric), f"{label}.allocation.{metric}")
        expect(metric_value.get("status") == "measured", f"{label}: {metric} status mismatch")
        expect(metric_value.get("scope") == ALLOCATOR_SCOPE, f"{label}: {metric} scope mismatch")
        values = [integer(item, f"{label}.allocation.{metric}.values[]") for item in constant_vector(metric_value.get("values"), samples, f"{label}.allocation.{metric}.values")]
        vectors[metric] = values
        expected_raw = [filesystem["allocation_by_index"][index][metric] for index in range(samples)]
        expect(values == expected_raw, f"{label}: {metric} operation vector differs from raw sample metrics")
    expect(vectors["failed_allocation_calls"] == [0] * samples, f"{label}: failed allocation calls are nonzero")
    return vectors


def validate_guard_result(report_info: dict[str, Any], label: str) -> dict[str, Any]:
    samples = report_info["sample_count"]
    report = report_info["_report"]
    expect(report.get("filesystem_evidence") == [], f"{label}: edit/save guard reports must not contain filesystem evidence")
    outputs: dict[str, Any] = {}
    for case in GUARD_CASES:
        result = report_info["results"][case]
        expect(result.get("output_sha256") == OUTPUT_SHA256, f"{label}.{case}: output hash mismatch")
        sink = obj(result.get("sink"), f"{label}.{case}.sink")
        expect(sink.get("accepted_bytes") == 4226480, f"{label}.{case}: sink byte count mismatch")
        expect(sink.get("write_calls") == 201, f"{label}.{case}: sink write count mismatch")
        expect(sink.get("largest_write") == 32768, f"{label}.{case}: sink largest write mismatch")
        outputs[case] = result["output_sha256"]
        if case == "xlsx_source_backed_cell_values_one_edit_save":
            source = obj(result.get("source"), f"{label}.{case}.source")
            cell_values = obj(source.get("xlsx_cell_values"), f"{label}.{case}.source.xlsx_cell_values")
            expect(cell_values.get("implementation") == "source-backed", f"{label}.{case}: source implementation mismatch")
            expect(cell_values.get("update_count") == 1, f"{label}.{case}: update count mismatch")
            expect(cell_values.get("selected_worksheet_count") == 1, f"{label}.{case}: selected worksheet count mismatch")
            expected_vectors = {
                "source_read_calls": 257,
                "source_read_bytes": 4233005,
                "workbook_read_calls": 1,
                "workbook_read_bytes": 226,
                "selected_worksheet_read_calls": 1,
                "selected_worksheet_read_bytes": 6816,
                "unselected_worksheet_read_calls": 3,
                "unselected_worksheet_read_bytes": 20330,
                "output_sha256": OUTPUT_SHA256,
                "semantic_sha256": OUTPUT_SEMANTIC_SHA256,
                "untouched_member_sha256": UNTOUCHED_MEMBER_SHA256,
            }
            for field, expected in expected_vectors.items():
                constant_vector(cell_values.get(field), samples, f"{label}.{case}.source.xlsx_cell_values.{field}", expected)
    expect(len(set(outputs.values())) == 1 and next(iter(outputs.values())) == OUTPUT_SHA256, f"{label}: eager/source output identities differ")
    return {"output_sha256": OUTPUT_SHA256, "output_semantic_sha256": OUTPUT_SEMANTIC_SHA256}


def validate_report_set(
    out: Path,
    kind: str,
    identities: dict[str, dict[str, Any]],
) -> tuple[dict[str, dict[str, Any]], dict[str, Any]]:
    reports: dict[str, dict[str, Any]] = {}
    raw_meta: dict[str, Any] = {}
    configurations: list[bytes] = []
    for leg in LEGS:
        path = report_path(out, leg, kind)
        value, info = load_json(path)
        report = obj(value, str(path))
        label = f"{leg}-{kind}"
        report_info = validate_report_header(report, label, kind, identities[role_for_leg(leg)])
        report_info["_report"] = report
        report_info["raw"] = info
        reports[leg] = report_info
        raw_meta[leg] = info
        configurations.append(canonical_bytes(report_info["configuration"]))
        catalog_value, _ = load_json(catalog_path(out, leg, kind))
        validate_catalog(catalog_value, f"{label}.catalog", expected_cases(kind), report_info["revision"])
    expect(len(set(configurations)) == 1, f"{kind}: configuration differs across ABBA legs")
    expect(len({info["sha256"] for info in raw_meta.values()}) == 4, f"{kind}: raw report files are not four distinct artifacts")
    return reports, {"reports": raw_meta, "configuration_sha256": sha256_bytes(configurations[0])}


def validate_commands(out: Path) -> dict[str, Any]:
    value, info = load_json(out / "commands.json")
    commands = list_value(value, "commands.json")
    by_label: dict[str, dict[str, Any]] = {}
    for item_value in commands:
        item = obj(item_value, "commands[]")
        label = item.get("label")
        expect(isinstance(label, str) and label not in by_label, f"commands.json: duplicate/invalid label {label!r}")
        by_label[label] = item
    expected: dict[str, dict[str, Any]] = {}
    for leg in LEGS:
        for kind in ALL_KINDS:
            label = f"{leg}-{kind}"
            command = obj(by_label.get(label), f"commands.json.{label}")
            expect(command.get("exit_code") == 0, f"{label}: command did not exit 0")
            argv = list_value(command.get("argv"), f"commands.json.{label}.argv")
            argv = [str(item) for item in argv]
            expect(argv[:5] == ["taskset", "-c", "2", "/usr/bin/time", "-v"], f"{label}: command prefix mismatch")
            expected_binary = "litchi-perf-baseline-alloc" if kind == "allocator" else "litchi-perf-baseline"
            expect(any(item.endswith("/" + expected_binary) or item == expected_binary for item in argv), f"{label}: executable mismatch")
            def option(name: str) -> str:
                try:
                    index = argv.index(name)
                    return argv[index + 1]
                except (ValueError, IndexError) as error:
                    raise VerificationError(f"{label}: missing {name}") from error
            expect(option("--case") == ",".join(expected_cases(kind)), f"{label}: case argv mismatch")
            expect(option("--filesystem-cache") == "warm", f"{label}: cache argv mismatch")
            expect(option("--xlsx-cell-crud-shape") == "medium", f"{label}: shape argv mismatch")
            samples, warmups = expected_samples(kind)
            expect(option("--samples") == str(samples), f"{label}: sample argv mismatch")
            expect(option("--warmup") == str(warmups), f"{label}: warmup argv mismatch")
            expected[label] = {"exit_code": command["exit_code"], "argv": argv, "cwd": command.get("cwd")}
    return {"journal": info, "commands": expected}


def metric_summary(vectors: dict[str, dict[str, list[int]]]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for metric in ALLOCATOR_METRICS:
        by_leg = {leg: vectors[leg][metric] for leg in LEGS}
        control_a1 = by_leg["a1"]
        control_a2 = by_leg["a2"]
        candidate_b1 = by_leg["b1"]
        candidate_b2 = by_leg["b2"]
        expect(control_a1 == control_a2, f"allocation {metric}: control A1/A2 vectors differ")
        expect(candidate_b1 == candidate_b2, f"allocation {metric}: candidate B1/B2 vectors differ")
        expect(len(set(control_a1)) == 1, f"allocation {metric}: control vector is not constant")
        expect(len(set(candidate_b1)) == 1, f"allocation {metric}: candidate vector is not constant")

        def stats(values: list[int]) -> dict[str, Any]:
            return {
                "count": len(values),
                "unique_values": sorted(set(values)),
                "min": min(values),
                "max": max(values),
                "mean": statistics.fmean(values),
            }

        control_value = statistics.fmean(control_a1)
        candidate_value = statistics.fmean(candidate_b1)
        def delta(control: float, candidate: float) -> dict[str, float | None]:
            absolute = candidate - control
            reduction = None if control == 0 else (control - candidate) / control * 100.0
            return {"candidate_minus_control": absolute, "candidate_reduction_percent": reduction}
        result[metric] = {
            "a1_control": stats(control_a1),
            "b1_candidate": stats(candidate_b1),
            "b2_candidate": stats(candidate_b2),
            "a2_control": stats(control_a2),
            "pair_deltas": {
                "a1_to_b1": delta(control_value, statistics.fmean(candidate_b1)),
                "a2_to_b2": delta(statistics.fmean(control_a2), statistics.fmean(candidate_b2)),
            },
            "combined": {
                "control": stats(control_a1 + control_a2),
                "candidate": stats(candidate_b1 + candidate_b2),
                "delta": delta(statistics.fmean(control_a1 + control_a2), statistics.fmean(candidate_b1 + candidate_b2)),
            },
        }
    return result


def verify(out: Path) -> dict[str, Any]:
    expect(out.is_dir(), f"capture directory does not exist: {out}")
    identities = {
        role: load_role_identity(out, role)[0]
        for role in ("control", "candidate")
    }
    command_info = validate_commands(out)
    selected_reports, selected_meta = validate_report_set(out, "normal", identities)
    guard_reports, guard_meta = validate_report_set(out, "guards", identities)
    allocator_reports, allocator_meta = validate_report_set(out, "allocator", identities)

    selected_children: set[int] = set()
    allocator_children: set[int] = set()
    selected_fs: dict[str, dict[str, Any]] = {}
    allocator_vectors: dict[str, dict[str, list[int]]] = {}
    for leg in LEGS:
        selected_fs[leg] = validate_selected_filesystem(
            selected_reports[leg], f"{leg}-normal", allocator=False, global_child_ids=selected_children
        )
        allocator_fs = validate_selected_filesystem(
            allocator_reports[leg], f"{leg}-allocator", allocator=True, global_child_ids=allocator_children
        )
        validate_guard_result(guard_reports[leg], f"{leg}-guards")
        allocator_vectors[leg] = validate_selected_result_allocation(
            allocator_reports[leg], f"{leg}-allocator", allocator_fs
        )

    allocation_summary = metric_summary(allocator_vectors)
    return {
        "schema": SCRIPT_VERSION,
        "claim_status": "allocator operation-scoped observations only; allocator elapsed latency is excluded",
        "capture": {
            "directory": str(out),
            "leg_order": ["a1_control", "b1_candidate", "b2_candidate", "a2_control"],
            "revisions": {"control": CONTROL_REVISION, "candidate": CANDIDATE_REVISION},
            "corpus": {
                "name": CORPUS["name"],
                "generator": CORPUS["generator"],
                "archive_sha256": ARCHIVE_SHA256,
                "archive_bytes": CORPUS["archive_bytes"],
                "input_semantic_sha256": INPUT_SEMANTIC_SHA256,
            },
            "selected_cell": SELECTED_CELL,
            "normal": {"samples_per_leg": 500, "warmups": 20, "reports": selected_meta},
            "guards": {"cases": list(GUARD_CASES), "samples_per_leg": 500, "warmups": 20, "reports": guard_meta},
            "allocator": {
                "samples_per_leg": 30,
                "warmups": 3,
                "reports": allocator_meta,
                "vectors_equal_to_raw_filesystem_samples": True,
                "control_a1_a2_vectors_equal": True,
                "candidate_b1_b2_vectors_equal": True,
            },
        },
        "binary_identity": {
            role: {
                "revision": identities[role]["revision"],
                "status": identities[role]["status"],
                "binaries": identities[role]["binaries"],
            }
            for role in ("control", "candidate")
        },
        "commands": command_info,
        "allocator_metric_summary": allocation_summary,
        "validation": {
            "normal_selected_child_process_ids": {leg: selected_fs[leg]["child_process_ids"] for leg in LEGS},
            "all_normal_selected_child_ids_unique": len(selected_children) == 4 * 500,
            "all_allocator_selected_child_ids_unique": len(allocator_children) == 4 * 30,
            "guard_output_sha256": OUTPUT_SHA256,
            "guard_output_semantic_sha256": OUTPUT_SEMANTIC_SHA256,
            "untouched_member_sha256": UNTOUCHED_MEMBER_SHA256,
            "allocator_elapsed_latency_compared": False,
        },
    }


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--out", "--root", dest="out", type=Path, default=DEFAULT_OUT,
        help="0410 capture directory (alias: --root)",
    )
    parser.add_argument(
        "--json-out", "--output", dest="json_out", type=Path,
        help="verification output (alias: --output; default: OUT/0410-verification.json)",
    )
    args = parser.parse_args(argv)
    output = args.json_out or args.out / "0410-verification.json"
    try:
        result = verify(args.out)
        output.write_text(json.dumps(result, indent=2, sort_keys=True, allow_nan=False) + "\n", encoding="utf-8")
    except (OSError, VerificationError, ValueError) as error:
        print(f"litchi-goal-0410-verify: FAIL: {error}", file=sys.stderr)
        return 2
    print(f"litchi-goal-0410-verify: PASS: {output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
