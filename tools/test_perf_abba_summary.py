import copy
import contextlib
import hashlib
import io
import json
import math
import tempfile
import unittest
from pathlib import Path

from tools import perf_abba_summary


TOOL = {
    "name": "litchi-perf-baseline",
    "version": "0.1.0",
    "binary": "litchi-perf-baseline",
    "profile": "release",
    "target_os": "linux",
    "target_arch": "x86_64",
    "instrumentation": "none",
}


CONFIGURATION = {
    "cases": ["synthetic_case"],
    "corpus_shapes": ["medium", "tiny"],
    "filesystem_root_selected": False,
    "samples_per_case": 15,
    "warmup_iterations_per_case": 1,
}


ENVIRONMENT = {
    "rustc_version": "rustc 1.95.0 (test)",
    "git_revision": "control-revision",
    "git_worktree_dirty": False,
    "logical_cpus_available": 1,
    "allocator": "Rust system allocator",
    "rustflags": None,
    "cargo_build_target": None,
    "perf_event_paranoid": "1",
}


def elapsed(samples):
    ordered = sorted(samples)
    mean = 0.0
    squared_deviation_sum = 0.0
    for index, value in enumerate(ordered):
        value_as_float = float(value)
        next_count = float(index + 1)
        delta = value_as_float - mean
        next_mean = mean + delta / next_count
        squared_deviation_sum += delta * (value_as_float - next_mean)
        mean = next_mean
    standard_deviation = math.sqrt(squared_deviation_sum / (len(ordered) - 1))
    critical_values = (
        12.706,
        4.303,
        3.182,
        2.776,
        2.571,
        2.447,
        2.365,
        2.306,
        2.262,
        2.228,
        2.201,
        2.179,
        2.160,
        2.145,
        2.131,
        2.120,
        2.110,
        2.101,
        2.093,
        2.086,
        2.080,
        2.074,
        2.069,
        2.064,
        2.060,
        2.056,
        2.052,
        2.048,
        2.045,
        2.042,
    )
    degrees_of_freedom = len(ordered) - 1
    if degrees_of_freedom == 0:
        critical = 0.0
    elif degrees_of_freedom <= len(critical_values):
        critical = critical_values[degrees_of_freedom - 1]
    else:
        z = 1.959963984540054
        z2 = z * z
        z3 = z2 * z
        z5 = z3 * z2
        z7 = z5 * z2
        degrees = float(degrees_of_freedom)
        critical = (
            z
            + (z3 + z) / (4.0 * degrees)
            + (5.0 * z5 + 16.0 * z3 + 3.0 * z)
            / (96.0 * degrees * degrees)
            + (3.0 * z7 + 19.0 * z5 + 17.0 * z3 - 15.0 * z)
            / (384.0 * degrees * degrees * degrees)
        )
    margin = critical * standard_deviation / math.sqrt(len(ordered))
    p50 = ordered[(len(ordered) - 1) // 2] // 2 + ordered[len(ordered) // 2] // 2
    p50 += (
        ordered[(len(ordered) - 1) // 2] % 2
        + ordered[len(ordered) // 2] % 2
    ) // 2
    return {
        "unit": "ns",
        "samples": ordered,
        "min": ordered[0],
        "p50": p50,
        "p95": ordered[max(1, (95 * len(samples) + 99) // 100) - 1],
        "p99": ordered[max(1, (99 * len(samples) + 99) // 100) - 1],
        "max": ordered[-1],
        "mean": mean,
        "standard_deviation": standard_deviation,
        "confidence_interval_95": {
            "method": "two-sided Student's t interval for the mean",
            "lower": max(mean - margin, 0.0),
            "upper": mean + margin,
        },
    }


def report(rows, *, revision="control-revision", dirty=False):
    environment = copy.deepcopy(ENVIRONMENT)
    environment["git_revision"] = revision
    environment["git_worktree_dirty"] = dirty
    candidate = revision.startswith("candidate")
    return {
        "schema_version": 1,
        "tool": copy.deepcopy(TOOL),
        "binary_identity": {
            "path": "/tmp/litchi-perf-candidate" if candidate else "/tmp/litchi-perf-control",
            "binary_sha256": "b" * 64 if candidate else "a" * 64,
            "binary_bytes": 200 if candidate else 100,
            "mode_bits": 0o755,
            "executable": True,
            "profile": "release",
        },
        "environment": environment,
        "configuration": copy.deepcopy(CONFIGURATION),
        "results": rows,
    }


def row(shape, samples, *, source=None, sink=None, output_sha256=None):
    samples = list(samples)
    if len(samples) < CONFIGURATION["samples_per_case"]:
        repeats = (CONFIGURATION["samples_per_case"] + len(samples) - 1) // len(samples)
        samples = (samples * repeats)[: CONFIGURATION["samples_per_case"]]
    result = {
        "case": "synthetic_case",
        "corpus": {
            "name": f"synthetic-{shape}",
            "shape": shape,
            "archive_sha256": shape * 64,
        },
        "elapsed_ns": elapsed(samples),
        "source": {"read_calls": [1] * len(samples)} if source is None else source,
        "sink": {"accepted_bytes": 10, "write_calls": 1} if sink is None else sink,
    }
    if output_sha256 is not None:
        result["output_sha256"] = output_sha256
    return result


def opc_multi_source(
    samples,
    mode,
    count,
    *,
    source_shape,
    payload_kind,
    uncompressed_payload_bytes,
    archive_bytes,
    archive_sha256,
    expected_eager_sha256,
    offset,
):
    samples = list(samples)
    output_digest = archive_sha256 if mode == "noop" else "c" * 64
    phase_preparation = [1] * len(samples)
    phase_open = [2] * len(samples)
    phase_planning = [1] * len(samples)
    phase_publication = [
        sample - phase_preparation[index] - phase_open[index] - phase_planning[index]
        for index, sample in enumerate(samples)
    ]
    overlay = {
        "implementation": "SourceBackedPackage::write_part_overlays_to_stream",
        "timing_scope": (
            "elapsed_ns is explicitly the sum of preparation_ns, open_ns, "
            "planning_ns, and publication_ns; interstitial work is excluded"
        ),
        "performance_claim": "none",
        "overlay_mode": mode,
        "replacement_semantics": {
            "changed": "non-empty changed-payload replacement plan",
            "noop": "non-empty equal-payload replacement plan; semantic no-op",
            "mixed": "non-empty mixed changed/equal-payload replacement plan",
        }[mode],
        "overlay_count": count,
        "source_shape": source_shape,
        "payload_kind": payload_kind,
        "source_bytes": archive_bytes,
        "source_sha256": archive_sha256,
        "expected_eager_sha256": expected_eager_sha256,
        "source_cache_max_bytes": uncompressed_payload_bytes,
        "source_cache_max_entries": 32,
        "sink_max_bytes": 68_004,
        "sink_max_write": 65_536,
        "preparation_ns": phase_preparation,
        "open_ns": phase_open,
        "planning_ns": phase_planning,
        "publication_ns": phase_publication,
        "cache_before_publication_hits": [0] * len(samples),
        "cache_before_publication_cold_loads": [0] * len(samples),
        "cache_before_publication_retained_entries": [0] * len(samples),
        "cache_before_publication_retained_bytes": [0] * len(samples),
        "source_cache_after_publication_probe_hits": [0] * len(samples),
        "source_cache_after_publication_probe_cold_loads": [count] * len(samples),
        "source_cache_after_publication_probe_retained_entries": [count] * len(samples),
        "source_cache_after_publication_probe_retained_bytes": [
            count * (uncompressed_payload_bytes // 32)
        ]
        * len(samples),
        "reopened_output_cache_hits": [0] * len(samples),
        "reopened_output_cache_cold_loads": [0] * len(samples),
        "reopened_output_cache_retained_entries": [0] * len(samples),
        "reopened_output_cache_retained_bytes": [0] * len(samples),
        "observed_after_publication_source_read_calls": [count + offset] * len(samples),
        "observed_after_publication_source_read_bytes": [10_000 + offset] * len(samples),
        "observed_after_publication_ordinary_payload_read_calls": [count] * len(samples),
        "observed_after_publication_ordinary_payload_read_bytes": [count * 123] * len(samples),
        "expected_eager_semantic_verified": True,
        "raw_members_and_order_preservation_verified": True,
        "equal_payload_noop_source_verified": mode == "noop",
        "observed_output_sha256": [output_digest] * len(samples),
    }
    return {
        "read_calls": [count + offset] * len(samples),
        "read_bytes": [10_000 + offset] * len(samples),
        "ordinary_payload_read_calls": [count] * len(samples),
        "ordinary_payload_read_bytes": [count * 123] * len(samples),
        "max_in_flight_reads": [1] * len(samples),
        "opc_source_overlay": overlay,
    }, output_digest


def opc_multi_sink():
    return {
        "accepted_bytes": 1234,
        "write_calls": 3,
        "largest_write": 512,
        "write_size_buckets": {
            "bytes_0": 0,
            "bytes_1_to_512": 3,
            "bytes_513_to_4096": 0,
            "bytes_4097_to_16384": 0,
            "bytes_16385_to_65536": 0,
            "bytes_over_65536": 0,
        },
    }


def docx_section_layout_sink():
    return {
        "accepted_bytes": 2_101_446,
        "write_calls": 105,
        "largest_write": 32_768,
        "write_size_buckets": {
            "bytes_0": 0,
            "bytes_1_to_512": 40,
            "bytes_513_to_4096": 1,
            "bytes_4097_to_16384": 0,
            "bytes_16385_to_65536": 64,
            "bytes_over_65536": 0,
        },
    }


def docx_section_layout_corpus():
    corpus = copy.deepcopy(perf_abba_summary.DOCX_SECTION_LAYOUT_CORPUS_IDENTITY)
    corpus.update(
        uncompressed_payload_bytes=2_116_441,
        archive_bytes=2_101_427,
        archive_sha256="183204a832a3518bf1d4474e9b7ff347ad7d0dc95ad2e9b1dbdde405837a29e8",
        target_payload_bytes=18_855,
        target_payload_sha256="67cb9066f112bfee6905a23f3eabdc76b65c86cd0c6613dd60a5a18895ceac81",
    )
    corpus["name"] = perf_abba_summary.DOCX_SECTION_LAYOUT_FIXED_CORPUS_FIELDS["name"]
    return corpus


def docx_section_layout_source(
    samples,
    *,
    offset=0,
    output_digest="5e87e9ca9fd6b9a98933c36d9aee1e848bc97eb7693229f7949514830121c4ec",
):
    samples = list(samples)
    preparation = [1] * len(samples)
    opened = [2] * len(samples)
    query = [1] * len(samples)
    edit = [1] * len(samples)
    commit = [1] * len(samples)
    publication = [
        sample
        - preparation[index]
        - opened[index]
        - query[index]
        - edit[index]
        - commit[index]
        for index, sample in enumerate(samples)
    ]
    phase_vectors = {
        "preparation_ns": preparation,
        "open_ns": opened,
        "query_ns": query,
        "edit_ns": edit,
        "commit_ns": commit,
        "publication_ns": publication,
    }
    read_calls = [3 + offset] * len(samples)
    read_bytes = [10_000 + offset] * len(samples)
    materializations = [1] * len(samples)
    overlay = {
        "implementation": "litchi-docx-source-backed-section-layout",
        "timing_scope": perf_abba_summary._DOCX_SECTION_LAYOUT_TIMING_SCOPE,
        "performance_claim": "none",
        "source_bytes": 2_101_427,
        "source_sha256": "183204a832a3518bf1d4474e9b7ff347ad7d0dc95ad2e9b1dbdde405837a29e8",
        "expected_output_sha256": output_digest,
        "total_main_story_paragraph_count": 257,
        "section_count": 3,
        "paragraph_owned_section_positions": [64, 129],
        "selected_paragraph": 64,
        "columns_before": 2,
        "columns_after": 3,
        "media_count": 4,
        "media_bytes": 512 * 1024,
        "sample_order": list(range(len(samples))),
        **phase_vectors,
        "source_read_calls": read_calls.copy(),
        "source_read_bytes": read_bytes.copy(),
        "ordinary_payload_materializations": materializations.copy(),
        "source_cache_before_publication_after_commit_hits": [0] * len(samples),
        "source_cache_before_publication_after_commit_cold_loads": [1] * len(samples),
        "source_cache_before_publication_after_commit_successful_loads": [1] * len(samples),
        "source_cache_before_publication_after_commit_retained_entries": [1] * len(samples),
        "source_cache_before_publication_after_commit_retained_bytes": [18_855]
        * len(samples),
        "output_sha256": [output_digest] * len(samples),
        "phase_sum_verified": True,
        "semantic_reopen_all_sections_verified": True,
        "table_cell_sectpr_excluded_verified": True,
        "raw_untouched_members_verified": True,
        "raw_member_order_and_comment_verified": True,
        "header_footer_relationships_verified": True,
        "source_immutability_verified": True,
        "patch_forward_inverse_verified": True,
        "publication_inverse_exact_source_verified": True,
        "stale_foreign_signed_noop_limits_partial_sink_verified": True,
        "cache_counters_verified": True,
        "sink_counters_verified": True,
    }
    return {
        "read_calls": read_calls.copy(),
        "read_bytes": read_bytes.copy(),
        "ordinary_payload_read_calls": [2] * len(samples),
        "ordinary_payload_read_bytes": [4_096] * len(samples),
        "max_in_flight_reads": [1] * len(samples),
        "ordinary_payload_materializations": materializations.copy(),
        "docx_section_layout": overlay,
    }


def docx_section_layout_row(
    samples,
    *,
    offset=0,
    output_digest="5e87e9ca9fd6b9a98933c36d9aee1e848bc97eb7693229f7949514830121c4ec",
):
    samples = list(samples)
    elapsed_ns = elapsed(samples)
    elapsed_ns["sample_order"] = list(range(len(samples)))
    return {
        "case": perf_abba_summary.DOCX_SECTION_LAYOUT_CASE,
        "corpus": docx_section_layout_corpus(),
        "elapsed_ns": elapsed_ns,
        "source": docx_section_layout_source(
            samples, offset=offset, output_digest=output_digest
        ),
        "sink": docx_section_layout_sink(),
        "output_sha256": output_digest,
    }


def reports_for_values(values):
    revisions = ("control-revision", "candidate-revision", "candidate-revision", "control-revision")
    return [
        report([row("tiny", samples)], revision=revision)
        for samples, revision in zip(values, revisions)
    ]


def four_legs():
    shapes = {
        "tiny": {
            "a1": [100, 100, 100, 100, 100],
            "b1": [80, 80, 80, 80, 80],
            "b2": [82, 82, 82, 82, 82],
            "a2": [102, 102, 102, 102, 102],
        },
        "medium": {
            "a1": [10, 20, 30, 40, 50],
            "b1": [8, 16, 24, 32, 40],
            "b2": [9, 18, 27, 36, 45],
            "a2": [11, 22, 33, 44, 55],
        },
    }
    revisions = {
        "a1": "control-revision",
        "b1": "candidate-revision",
        "b2": "candidate-revision",
        "a2": "control-revision",
    }
    return [
        report(
            [row(shape, values[label]) for shape, values in shapes.items()],
            revision=revisions[label],
        )
        for label in ("a1", "b1", "b2", "a2")
    ]


def xls_operation_item(
    mode="generic_read_at",
    source_bytes=131072,
    implementation="source_backed",
    family="Number",
):
    corpus_contract = perf_abba_summary.XLS_NUMERIC_CORPUS_CONTRACTS[family]
    output_contract = perf_abba_summary.XLS_NUMERIC_OUTPUT_CONTRACTS[family]
    fingerprint_chunks = (source_bytes + 1_048_576 - 1) // 1_048_576
    publication_chunks = (source_bytes + 65_536 - 1) // 65_536
    fingerprint_bytes = source_bytes * 2
    fenced = 1 if mode == "generic_read_at" else 0

    def phase(scans, emission=False):
        return {
            "scans": scans,
            "bytes": fingerprint_bytes * scans,
            "chunks": (publication_chunks if emission else fingerprint_chunks) * scans,
        }

    materialized = implementation != "plan_only"
    update_count = 1 if family == "Number" else 3
    replacement_bytes = output_contract["replacement_bytes"]
    operation = {
        "counter_scope": perf_abba_summary.XLS_NUMERIC_OPERATION_COUNTER_SCOPE,
        "source_mode": mode,
        "source_bytes": source_bytes,
        "fingerprint_chunk_bytes": 1_048_576,
        "publication_chunk_bytes": 65_536,
        "fingerprint_buffer_bytes": min(source_bytes, 1_048_576),
        "publication_buffer_bytes": 65_536,
        "planning_fingerprint_scans": 1 + fenced,
        "planning_fingerprint_bytes": fingerprint_bytes * (1 + fenced),
        "planning_fingerprint_chunks": fingerprint_chunks * (1 + fenced),
        "composed_source_preflight_scans": 1,
        "composed_source_preflight_bytes": fingerprint_bytes,
        "composed_source_preflight_chunks": fingerprint_chunks,
        "candidate_reopen_logical_artifact_bytes": source_bytes if not materialized else source_bytes,
        "selected_stream_logical_bytes": corpus_contract["target_payload_bytes"],
        "splice_count": update_count,
        "changed_span_count": output_contract["changed_spans"],
        "replacement_bytes": replacement_bytes,
        "target_materialization_vec_bytes": source_bytes if materialized else 0,
        "target_materialization_clone_bytes": source_bytes if materialized else 0,
        "publication_write_calls": output_contract["sink_write_calls"],
        "atomic_save_event_scope": perf_abba_summary.XLS_NUMERIC_OPERATION_ATOMIC_SCOPE,
    }
    for prefix in (
        "target_materialization_write_pre",
        "target_materialization_write_post",
        "direct_write_pre",
        "direct_write_post",
        "atomic_save_pre_temp",
        "atomic_save_pre_rename",
    ):
        values = phase(fenced)
        operation.update({f"{prefix}_{key}": value for key, value in values.items()})
    for prefix in (
        "target_materialization_emission",
        "direct_emission",
        "atomic_save_emission",
    ):
        values = phase(1, emission=True)
        operation.update({f"{prefix}_{key}": value for key, value in values.items()})
    return operation


def xls_numeric_source(
    mode="generic_read_at",
    implementation="source_backed",
    timing_offset=0,
    include_operation=True,
    family="Number",
):
    corpus_contract = perf_abba_summary.XLS_NUMERIC_CORPUS_CONTRACTS[family]
    workbook_sizes = perf_abba_summary.XLS_NUMERIC_WORKBOOK_SIZE_CONTRACTS[family]
    output_contract = perf_abba_summary.XLS_NUMERIC_OUTPUT_CONTRACTS[family]
    source_bytes = corpus_contract["archive_bytes"]
    sink_write_calls = output_contract["sink_write_calls"]
    source_read_calls = output_contract["source_read_calls"]
    source_read_bytes = output_contract["source_read_bytes"]
    materialized = implementation != "plan_only"
    update_count = 1 if family == "Number" else 3
    numeric = {
        "source_counter_scope": (
            perf_abba_summary.XLS_NUMERIC_CURRENT_SOURCE_COUNTER_SCOPE
            if implementation != "eager"
            else perf_abba_summary.XLS_NUMERIC_LEGACY_SOURCE_COUNTER_SCOPE
        ),
        "implementation": implementation,
        "family": family,
        "source_backed": implementation != "eager",
        "target_artifact_retained_at_commit": materialized,
        "target_artifact_materialized_at_commit": materialized,
        "patch_or_inverse_supported": materialized,
        "update_count": update_count,
        "sample_count": 15,
        "input_cfb_bytes": source_bytes,
        "output_cfb_bytes": source_bytes,
        "source_workbook_bytes": workbook_sizes["source_workbook_bytes"],
        "target_workbook_bytes": workbook_sizes["target_workbook_bytes"],
        "sink_capacity_bytes": source_bytes,
        "expected_output_sha256": output_contract["output_sha256"],
        "owned_input_scope": "complete in-memory CFB bytes; no positional/physical I/O",
        "edit_ns": [100 + timing_offset + index for index in range(15)],
        "set_ns": [200 + timing_offset + index for index in range(15)],
        "commit_ns": [300 + timing_offset + index for index in range(15)],
        "publication_ns": [400 + timing_offset + index for index in range(15)],
        "total_ns": [1000 + 4 * timing_offset + 4 * index for index in range(15)],
        "complete_target_materialized_bytes": [source_bytes] * 15,
        "sink_bytes": [source_bytes] * 15,
        "sink_write_calls": [sink_write_calls] * 15,
        "sink_digests": [output_contract["target_fingerprint_sha256"]] * 15,
        "source_bytes": [source_bytes] * 15,
        "source_workbook_bytes_per_sample": [
            workbook_sizes["source_workbook_bytes"]
        ]
        * 15,
        "target_workbook_bytes_per_sample": [
            workbook_sizes["target_workbook_bytes"]
        ]
        * 15,
        "splice_count": [update_count] * 15,
        "replacement_bytes": [output_contract["replacement_bytes"]] * 15,
        "changed_spans": [output_contract["changed_spans"]] * 15,
        "source_fingerprints": [output_contract["source_fingerprint_sha256"]] * 15,
        "target_fingerprints": [output_contract["target_fingerprint_sha256"]] * 15,
    }
    if not materialized:
        numeric["complete_target_materialized_bytes"] = [0] * 15
    if implementation == "eager":
        for field in (
            "splice_count",
            "replacement_bytes",
            "changed_spans",
            "source_fingerprints",
            "target_fingerprints",
        ):
            numeric.pop(field)
    elif include_operation:
        operation = xls_operation_item(mode, source_bytes, implementation, family)
        numeric["operation_evidence_schema"] = perf_abba_summary.XLS_NUMERIC_OPERATION_EVIDENCE_SCHEMA
        numeric["operation_evidence"] = [copy.deepcopy(operation) for _ in range(15)]
    return {
        "read_calls": [source_read_calls] * 15,
        "read_bytes": [source_read_bytes] * 15,
        "xls_numeric": numeric,
    }


def xls_numeric_legs(
    case="xls_numeric_source_backed_number_edit_save",
    implementation="source_backed",
    family="Number",
):
    legs = four_legs()
    modes = (
        ("owned_immutable_arc",) * 4
        if implementation == "plan_only"
        else ("generic_read_at", "owned_immutable_arc", "owned_immutable_arc", "generic_read_at")
    )
    for index, (leg, mode) in enumerate(zip(legs, modes)):
        leg["results"] = [leg["results"][0]]
        leg["configuration"]["cases"] = [case]
        for result in leg["results"]:
            result["case"] = case
        result = leg["results"][0]
        result["corpus"] = copy.deepcopy(
            perf_abba_summary.XLS_NUMERIC_CORPUS_CONTRACTS[family]
        )
        leg["configuration"]["corpus_shapes"] = [
            "tiny",
            "many-small",
            "few-large",
            "wide-root",
        ]
        result["output_sha256"] = perf_abba_summary.XLS_NUMERIC_OUTPUT_CONTRACTS[family][
            "output_sha256"
        ]
        result["sink"] = {
            "accepted_bytes": result["corpus"]["archive_bytes"],
            "write_calls": perf_abba_summary.XLS_NUMERIC_OUTPUT_CONTRACTS[family][
                "sink_write_calls"
            ],
        }
        result["source"] = xls_numeric_source(
            mode,
            implementation=implementation,
            timing_offset=index,
            family=family,
        )
        result["elapsed_ns"] = elapsed(
            result["source"]["xls_numeric"]["total_ns"]
        )
        result["elapsed_ns"]["sample_order"] = list(range(15))
    return legs


def xls_numeric_all_selector_legs():
    selectors = tuple(perf_abba_summary.XLS_NUMERIC_CASE_CONTRACTS.items())
    legs = four_legs()
    modes = (
        "generic_read_at",
        "owned_immutable_arc",
        "owned_immutable_arc",
        "generic_read_at",
    )
    for leg, mode, leg_index in zip(legs, modes, range(4)):
        results = []
        for case, (implementation, family, _update_count) in selectors:
            result = copy.deepcopy(leg["results"][0])
            result["case"] = case
            result["corpus"] = copy.deepcopy(
                perf_abba_summary.XLS_NUMERIC_CORPUS_CONTRACTS[family]
            )
            result["output_sha256"] = perf_abba_summary.XLS_NUMERIC_OUTPUT_CONTRACTS[
                family
            ]["output_sha256"]
            result["sink"] = {
                "accepted_bytes": result["corpus"]["archive_bytes"],
                "write_calls": perf_abba_summary.XLS_NUMERIC_OUTPUT_CONTRACTS[family][
                    "sink_write_calls"
                ],
            }
            source_mode = "owned_immutable_arc" if implementation == "plan_only" else mode
            result["source"] = xls_numeric_source(
                source_mode,
                implementation=implementation,
                timing_offset=leg_index,
                family=family,
            )
            result["elapsed_ns"] = elapsed(
                result["source"]["xls_numeric"]["total_ns"]
            )
            result["elapsed_ns"]["sample_order"] = list(range(15))
            results.append(result)
        leg["results"] = results
        leg["configuration"]["cases"] = [case for case, _ in selectors]
        leg["configuration"]["corpus_shapes"] = [
            "tiny",
            "many-small",
            "few-large",
            "wide-root",
        ]
    return legs


def with_parallel_metrics(reports):
    reports = copy.deepcopy(reports)
    for report_value in reports:
        report_value["configuration"]["execution_workers"] = [1, 2]
        cases = []
        for result in report_value["results"]:
            samples = result["elapsed_ns"]["samples"]
            sample_order = list(range(len(samples)))
            result["elapsed_ns"]["sample_order"] = sample_order
            result["execution"] = {"worker_count": 2, "logical_tasks": 3}
            result["source"]["simulation"] = {
                "physical_request_count": list(range(1, len(samples) + 1))
            }
            cases.append(
                {
                    "case": result["case"],
                    "corpus_sha256": result["corpus"]["archive_sha256"],
                    "configured_worker_count": {
                        "status": "measured",
                        "value": 2,
                        "scope": "result.execution.worker_count",
                    },
                    "observed_local_worker_count": {
                        "status": "not_applicable",
                        "scope": "result.source.opc_cache.worker_count_with_one_"
                        "created_local_worker_team",
                        "reason": "result does not create an explicit local worker team",
                    },
                    "deterministic_task_count": {
                        "status": "measured",
                        "value": 3,
                        "scope": "result.execution.logical_tasks",
                    },
                    "deterministic_chunk_count": {
                        "status": "measured",
                        "value": list(range(1, len(samples) + 1)),
                        "scope": "result.source.simulation.physical_request_count",
                    },
                    "lock_wait_ns": {
                        "status": "unavailable",
                        "scope": "lock_wait_ns",
                        "reason": "no exact instrumented lock boundary is present",
                    },
                }
            )
        report_value["parallel_metrics"] = {
            "schema_version": 1,
            "scope": "explicit_local_execution_only",
            "claim": "descriptive",
            "configured_worker_budget": {
                "status": "measured",
                "value": [1, 2],
                "scope": "configuration.execution_workers",
            },
            "observed_process_thread_count": {
                "status": "unavailable",
                "scope": "process_thread_count",
                "reason": "no process-global thread counter is collected",
            },
            "cases": cases,
        }
    return reports


def with_operation_metrics(reports):
    """Attach the current strict operation-metrics schema to every result."""

    from tools.test_perf_compare import operation_metrics_report_fields

    def resize_vectors(value):
        if isinstance(value, dict):
            for key, item in value.items():
                if key == "values" and isinstance(item, list):
                    value[key] = (item * 4)[:15]
                else:
                    resize_vectors(item)

    reports = copy.deepcopy(reports)
    for leg in reports:
        for result in leg["results"]:
            metrics = operation_metrics_report_fields()
            metrics["sample_count"] = 15
            metrics["sample_indices"] = list(range(15))
            resize_vectors(metrics)
            result["operation_metrics"] = metrics
    return reports


def with_xlsx_operation_metrics(reports):
    """Attach valid filesystem-XLSX operation metrics and sample identity."""

    reports = with_operation_metrics(reports)
    for leg in reports:
        for result in leg["results"]:
            sample_count = len(result["elapsed_ns"]["samples"])
            sample_order = list(range(sample_count))
            result["elapsed_ns"]["sample_order"] = sample_order
            metrics = result["operation_metrics"]
            metrics["sample_count"] = sample_count
            metrics["sample_indices"] = sample_order
            source = metrics["source"]
            source["status"] = "not_applicable"
            source["counter_scope"] = "not_applicable_filesystem_xlsx"
            for field, vector in source.items():
                if field in {"status", "counter_scope"}:
                    continue
                vector["status"] = "not_applicable"
                vector.pop("values", None)
    return reports


def with_legacy_operation_metrics(reports):
    """Attach the pre-additive schema-1 operation-metrics envelope."""

    reports = with_operation_metrics(reports)
    for leg in reports:
        for result in leg["results"]:
            current = result["operation_metrics"]
            result["operation_metrics"] = {
                "sample_count": current["sample_count"],
                "alignment": "elapsed_ns.samples",
                "source": {
                    key: current["source"][key]
                    for key in (
                        "status",
                        "counter_scope",
                        "logical_read_calls",
                        "logical_read_requested_bytes",
                        "logical_read_returned_bytes",
                        "max_concurrent_reads",
                    )
                },
                "process": {
                    key: current["process"][key]
                    for key in (
                        "status",
                        "user_cpu_ticks",
                        "system_cpu_ticks",
                        "clock_ticks_per_second",
                        "minor_faults",
                        "major_faults",
                        "voluntary_context_switches",
                        "nonvoluntary_context_switches",
                        "rss_delta_bytes",
                        "peak_rss_bytes",
                    )
                },
                "sink": {
                    key: current["sink"][key] for key in ("status", "output_bytes")
                },
                "publication": current["publication"],
                "materialization": current["materialization"],
                "cfb_phases": current["cfb_phases"],
            }
    return reports


def with_filesystem_evidence(reports):
    reports = copy.deepcopy(reports)
    for leg in reports:
        leg["configuration"]["filesystem_cache_states"] = ["warm"]
        leg["configuration"]["filesystem_fresh_child_per_sample"] = True
        result = leg["results"][0]
        leg["filesystem_evidence"] = [
            {
                "case": result["case"],
                "corpus": copy.deepcopy(result["corpus"]),
                "warmup_iterations": 1,
                "sample_count": 15,
                "cache_states": ["warm"],
                "fresh_child_per_sample": True,
                "samples": [
                    {
                        "sample_index": index,
                        "cache_state": "warm",
                        "parent_wall_ns": 200 + index,
                        "cold_advice": "not_requested",
                        "logical_read_counter_scope": "test_scope",
                        "logical_read_calls": 0,
                        "logical_read_requested_bytes": 0,
                        "logical_read_bytes": 0,
                        "logical_read_largest_requested_bytes": 0,
                        "logical_read_largest_returned_bytes": 0,
                        "max_concurrent_reads": 0,
                        "logical_read_request_sizes": [],
                        "logical_read_request_size_buckets": {
                            "bytes_0": 0,
                            "bytes_1_to_512": 0,
                            "bytes_513_to_4096": 0,
                            "bytes_4097_to_16384": 0,
                            "bytes_16385_to_65536": 0,
                            "bytes_over_65536": 0,
                        },
                        "process_metrics": None,
                        "output_sha256": None,
                        "output_bytes": None,
                        "opc_materialized_parts": None,
                        "cfb_changed_spans": None,
                        "cfb_published_bytes": None,
                        "elapsed_ns": 100 + index,
                    }
                    for index in range(15)
                ],
                "tool": copy.deepcopy(leg["tool"]),
                "configuration": copy.deepcopy(leg["configuration"]),
            }
        ]
    return reports


def with_xlsx_repeat_store_evidence(
    *,
    structural=False,
    scenario="medium",
    child_process_ids=True,
    pid_offset=0,
):
    """Build four compact repeated-store filesystem reports for verifier tests."""

    case = f"xlsx_source_repeated_store_{scenario}"
    if structural:
        case += "_reacquisition_control"
    source_sha256 = (
        "3cf797e44ef51189a4b62d040cf39ff2af670ebd909c6e806f387b51e72ecfec"
        if scenario == "oversized"
        else "dfff7ec0c749d9e404091776f15a8fb690985af7f58efdfe659dbeaed7145036"
    )
    full_semantic_sha256 = (
        "020fdd140d2959ea4f480676a3d4d0bf840927e25251cb6cad37a043ab80627e"
    )
    semantic_projection_sha256 = (
        "01c253bf3fc611835e0806414c6417a9cfbb012ff6e01f9bb55cec94236a6235"
    )
    archive_bytes = 4_236_114 if scenario == "oversized" else 4_226_429
    memory_limit = archive_bytes * 4 + 64 * 1024 * 1024
    selected_member_bytes = (
        8_389_041 if scenario == "oversized" else 63_294
    )
    query_count = 8 * 4
    revisions = (
        "control-revision",
        "candidate-revision",
        "candidate-revision",
        "control-revision",
    )
    labels = ("a1", "b1", "b2", "a2")

    def counters(**overrides):
        value = {
            "cache_cold_loads": 0,
            "cache_successful_loads": 0,
            "cache_bypasses": 0,
            "cache_oversized_bypasses": 0,
            "cache_evictions": 0,
            "source_read_calls": 0,
            "source_read_bytes": 0,
            "selected_member_read_calls": 0,
            "selected_member_read_bytes": 0,
            "budget_input_bytes_used": 0,
            "budget_work_used": 0,
        }
        value.update(overrides)
        return value

    def repeat_for(label):
        query_elapsed = [result_elapsed // 4] * 3
        query_elapsed.append(result_elapsed - sum(query_elapsed))
        if structural:
            if scenario == "medium":
                delta = counters(
                    cache_cold_loads=96,
                    cache_successful_loads=96,
                    cache_evictions=96,
                    source_read_calls=query_count,
                    source_read_bytes=3_200,
                    selected_member_read_calls=query_count,
                    selected_member_read_bytes=3_200,
                    budget_input_bytes_used=3_200,
                    budget_work_used=96,
                )
                before = counters(
                    cache_cold_loads=1,
                    cache_successful_loads=1,
                    cache_evictions=1,
                    budget_work_used=1,
                )
            else:
                delta = counters(
                    cache_cold_loads=32,
                    cache_successful_loads=32,
                    cache_bypasses=32,
                    cache_oversized_bypasses=32,
                    source_read_calls=query_count,
                    source_read_bytes=3_200,
                    selected_member_read_calls=query_count,
                    selected_member_read_bytes=3_200,
                    budget_input_bytes_used=3_200,
                    budget_work_used=32,
                )
                before = counters(
                    cache_successful_loads=1,
                    cache_bypasses=1,
                    cache_oversized_bypasses=1,
                )
            implementation = (
                "explicit_part_data_reacquisition_structural_control"
            )
            claim_scope = (
                "structural cache/read control only; elapsed/query_ns must not be "
                "compared with candidate"
            )
            control_reacquire_count = query_count
        else:
            if scenario == "medium":
                before = counters(
                    cache_cold_loads=1,
                    cache_successful_loads=1,
                    cache_evictions=1,
                    budget_work_used=1,
                    source_read_calls=1,
                    source_read_bytes=100,
                    selected_member_read_calls=1,
                    selected_member_read_bytes=100,
                )
            else:
                before = counters(
                    cache_successful_loads=1,
                    cache_bypasses=1,
                    cache_oversized_bypasses=1,
                    source_read_calls=1,
                    source_read_bytes=100,
                    selected_member_read_calls=1,
                    selected_member_read_bytes=100,
                )
            delta = (
                counters(
                    cache_cold_loads=1,
                    cache_successful_loads=1,
                    cache_evictions=1,
                    source_read_calls=1,
                    source_read_bytes=100,
                    selected_member_read_calls=1,
                    selected_member_read_bytes=100,
                    budget_input_bytes_used=100,
                    budget_work_used=1,
                )
                if label in {"a1", "a2"}
                else counters()
            )
            implementation = "source_backed_cached_store"
            claim_scope = (
                "primary repeated-query selector; compare only the same selector "
                "across A/B revisions"
            )
            control_reacquire_count = 0
        after = {
            key: before[key] + delta[key] for key in before
        }
        return {
            "implementation": implementation,
            "scenario": scenario,
            "selected_member": "xl/worksheets/sheet1.xml",
            "selected_member_uncompressed_bytes": selected_member_bytes,
            "cache_max_bytes": 8 * 1024 * 1024,
            "cache_max_entries": 2 if scenario == "medium" else 128,
            "query_iterations": 8,
            "query_names": ["cell", "cells", "visit", "stored_extent"],
            "query_elapsed_ns": query_elapsed,
            "timed_elapsed_total_ns": result_elapsed,
            "control_reacquire_count": control_reacquire_count,
            "timing_scope": "semantic_query_only; explicit PartData reacquisition excluded",
            "claim_scope": claim_scope,
            "budget_managed": True,
            "budget_memory_limit": memory_limit,
            "budget_input_bytes_limit": (1 << 64) - 1,
            "budget_work_limit": (1 << 64) - 1,
            "semantic_projection_sha256": semantic_projection_sha256,
            "diagnostics_before": before,
            "diagnostics_after": after,
            "diagnostics_delta": delta,
        }

    reports = []
    for leg_index, (label, revision) in enumerate(zip(labels, revisions)):
        result_elapsed = (
            80
            if label == "b1"
            else 82
            if label == "b2"
            else 100
            if label == "a1"
            else 102
        )
        result = row(case, [result_elapsed], source=None, sink=None)
        result["source"] = None
        result["sink"] = None
        result["output_sha256"] = None
        result["elapsed_ns"]["sample_order"] = list(range(15))
        result["case"] = case
        result["corpus"] = (
            {
                "name": "xlsx-source-repeated-store-oversized",
                "generator": "litchi-xlsx-source-repeated-store-corpus-v1",
                "package_format": "XLSX/OPC/ZIP",
                "shape": "oversized",
                "payload_kind": "fixed-medium-grid-with-oversized-selected-worksheet",
                "compression": "deflate",
                "entry_count": 9216,
                "archive_member_count": 17,
                "entry_bytes": 4,
                "uncompressed_payload_bytes": 12_789_836,
                "archive_bytes": 4_236_114,
                "archive_sha256": source_sha256,
                "target_entry": "Sheet1!A1",
                "target_payload_bytes": 1,
                "target_payload_sha256": "5feceb66ffc86f38d952786c6d696c79c2dbc239dd4e91b46729d73a27fb57e9",
                "xlsx": {
                    "sheet_count": 4,
                    "rows_per_sheet": 48,
                    "columns_per_sheet": 48,
                    "one_percent_update_count": 93,
                    "source_members": {
                        "workbook": "xl/workbook.xml",
                        "worksheets": [
                            "xl/worksheets/sheet1.xml",
                            "xl/worksheets/sheet2.xml",
                            "xl/worksheets/sheet3.xml",
                            "xl/worksheets/sheet4.xml",
                        ],
                        "shared_strings": None,
                        "styles": "xl/styles.xml",
                    },
                },
            }
            if scenario == "oversized"
            else {
                "name": "xlsx-source-repeated-store-medium",
                "generator": "litchi-xlsx-source-repeated-store-corpus-v1",
                "package_format": "XLSX/OPC/ZIP",
                "shape": "medium",
                "payload_kind": "fixed-medium-grid-for-repeated-selected-store",
                "compression": "deflate",
                "entry_count": 9216,
                "archive_member_count": 17,
                "entry_bytes": 4,
                "uncompressed_payload_bytes": 4_231_168,
                "archive_bytes": 4_226_429,
                "archive_sha256": source_sha256,
                "target_entry": "Sheet1!A1",
                "target_payload_bytes": 1,
                "target_payload_sha256": "5feceb66ffc86f38d952786c6d696c79c2dbc239dd4e91b46729d73a27fb57e9",
                "xlsx": {
                    "sheet_count": 4,
                    "rows_per_sheet": 48,
                    "columns_per_sheet": 48,
                    "one_percent_update_count": 93,
                    "source_members": {
                        "workbook": "xl/workbook.xml",
                        "worksheets": [
                            "xl/worksheets/sheet1.xml",
                            "xl/worksheets/sheet2.xml",
                            "xl/worksheets/sheet3.xml",
                            "xl/worksheets/sheet4.xml",
                        ],
                        "shared_strings": None,
                        "styles": "xl/styles.xml",
                    },
                },
            }
        )
        leg = report([result], revision=revision)
        leg["configuration"]["cases"] = [case]
        leg["configuration"]["corpus_shapes"] = [scenario]
        leg["configuration"]["filesystem_cache_states"] = ["warm"]
        leg["configuration"]["filesystem_fresh_child_per_sample"] = True
        evidence = {
            "case": case,
            "corpus": copy.deepcopy(result["corpus"]),
            "warmup_iterations": 1,
            "sample_count": 15,
            "cache_states": ["warm"],
            "fresh_child_per_sample": True,
            "samples": [],
            "tool": copy.deepcopy(leg["tool"]),
            "configuration": copy.deepcopy(leg["configuration"]),
        }
        for sample_index in range(15):
            sample = {
                "sample_index": sample_index,
                "cache_state": "warm",
                "elapsed_ns": result_elapsed,
                "parent_wall_ns": 200 + sample_index,
                "cold_advice": "not_requested",
                "logical_read_counter_scope": "not_applicable_filesystem_xlsx",
                "logical_read_calls": 0,
                "logical_read_requested_bytes": 0,
                "logical_read_bytes": 0,
                "logical_read_largest_requested_bytes": 0,
                "logical_read_largest_returned_bytes": 0,
                "max_concurrent_reads": 0,
                "logical_read_request_sizes": [],
                "logical_read_request_size_buckets": {
                    "bytes_0": 0,
                    "bytes_1_to_512": 0,
                    "bytes_513_to_4096": 0,
                    "bytes_4097_to_16384": 0,
                    "bytes_16385_to_65536": 0,
                    "bytes_over_65536": 0,
                },
                "process_metrics": None,
                "output_sha256": None,
                "output_bytes": None,
                "opc_materialized_parts": None,
                "cfb_changed_spans": None,
                "cfb_published_bytes": None,
                "xlsx_source_sha256": source_sha256,
                "xlsx_semantic_sha256": full_semantic_sha256,
                "xlsx_repeat_store": repeat_for(label),
            }
            if child_process_ids:
                sample["child_process_id"] = (
                    1_000 + pid_offset + leg_index * 100 + sample_index
                )
            evidence["samples"].append(sample)
        leg["filesystem_evidence"] = [evidence]
        reports.append(leg)
    return reports


class PerfAbbaSummaryTests(unittest.TestCase):
    def test_recomputes_statistics_and_emits_every_multi_shape_row(self):
        summary = perf_abba_summary.summarize_reports(four_legs())
        self.assertEqual(summary["verification"]["result_count"], 2)
        self.assertEqual([item["shape"] for item in summary["results"]], ["medium", "tiny"])
        medium = summary["results"][0]
        elapsed_summary = medium["elapsed_ns"]
        self.assertEqual(elapsed_summary["legs_ns"]["a1"]["p50"], 30)
        self.assertEqual(elapsed_summary["legs_ns"]["a1"]["mean"], 30.0)
        self.assertEqual(elapsed_summary["legs_ns"]["a1"]["p95"], 50)
        self.assertEqual(elapsed_summary["legs_ns"]["a1"]["p99"], 50)
        self.assertAlmostEqual(
            elapsed_summary["candidate_reduction_percent"]["a1_to_b1"]["mean"], 20.0
        )
        self.assertAlmostEqual(
            elapsed_summary["same_implementation_drift_percent"]["control"]["p50"], 10.0
        )
        self.assertEqual(elapsed_summary["accepted_statistics"], ["p99"])

    def test_validates_descriptive_parallel_metrics_without_comparing_them(self):
        reports = with_parallel_metrics(four_legs())
        summary = perf_abba_summary.summarize_reports(reports)
        self.assertEqual(summary["verification"]["result_count"], 2)

        combined = with_operation_metrics(with_parallel_metrics(four_legs()))
        combined_summary = perf_abba_summary.summarize_reports(combined)
        self.assertEqual(
            combined_summary["results"][0]["identity"]["operation_metrics_status"],
            "verified_equal",
        )

        malformed = copy.deepcopy(reports)
        malformed[0]["parallel_metrics"]["cases"][0]["deterministic_task_count"][
            "value"
        ] = [3]
        with self.assertRaisesRegex(
            perf_abba_summary.AbbaSummaryInputError,
            "deterministic_task_count.value must be a non-negative",
        ):
            perf_abba_summary.summarize_reports(malformed)

    def test_operation_metrics_validate_nested_identity_but_not_metric_values(self):
        summary = perf_abba_summary.summarize_reports(with_operation_metrics(four_legs()))
        self.assertEqual(
            summary["results"][0]["identity"]["operation_metrics_status"],
            "verified_equal",
        )

        mutations = (
            (
                lambda legs: legs[1]["results"][0]["operation_metrics"].update(
                    sample_count=14
                ),
                "operation_metrics.*sample_count",
            ),
            (
                lambda legs: legs[1]["results"][0]["operation_metrics"].update(
                    schema=2
                ),
                "operation_metrics keys mismatch",
            ),
            (
                lambda legs: legs[1]["results"][0]["operation_metrics"]["source"].update(
                    counter_scope="untimed_source_replay_only"
                ),
                "source.status=.*incompatible",
            ),
        )
        for mutation, message in mutations:
            legs = with_operation_metrics(four_legs())
            mutation(legs)
            with self.subTest(message=message), self.assertRaisesRegex(
                perf_abba_summary.AbbaSummaryInputError, message
            ):
                perf_abba_summary.summarize_reports(legs)

        numeric_change = with_operation_metrics(four_legs())
        numeric_change[1]["results"][0]["operation_metrics"]["process"]["read_bytes"][
            "values"
        ] = [999] * 15
        self.assertEqual(
            perf_abba_summary.summarize_reports(numeric_change)["results"][0][
                "identity"
            ]["operation_metrics_status"],
            "verified_equal",
        )

    def test_operation_metrics_identity_ignores_valid_cross_leg_sample_order(self):
        reports = with_xlsx_operation_metrics(four_legs())
        sample_indices = reports[1]["results"][1]["operation_metrics"][
            "sample_indices"
        ]
        reordered_sample_indices = [
            *sample_indices[3:6],
            *sample_indices[:3],
            *sample_indices[6:],
        ]
        reports[1]["results"][1]["operation_metrics"][
            "sample_indices"
        ] = reordered_sample_indices
        reports[1]["results"][1]["elapsed_ns"][
            "sample_order"
        ] = reordered_sample_indices
        self.assertEqual(
            perf_abba_summary._operation_metrics_identity_projection(
                reports[1]["results"][1]["operation_metrics"]
            )["sample_indices"],
            "<permutation>",
        )

        summary = perf_abba_summary.summarize_reports(reports)
        self.assertEqual(
            summary["results"][0]["identity"]["operation_metrics_status"],
            "verified_equal",
        )

    def test_operation_metrics_identity_preserves_non_xlsx_sample_indices(self):
        reports = with_operation_metrics(four_legs())
        operation_metrics = reports[0]["results"][0]["operation_metrics"]
        operation_metrics["source"]["status"] = "not_applicable"
        operation_metrics["source"][
            "counter_scope"
        ] = "not_applicable_in_process_sink"
        projected = perf_abba_summary._operation_metrics_identity_projection(
            operation_metrics
        )
        self.assertEqual(projected["sample_indices"], list(range(15)))

    def test_legacy_operation_metrics_use_exact_historical_schema(self):
        reports = with_legacy_operation_metrics(four_legs())
        summary = perf_abba_summary.summarize_reports(reports)
        self.assertEqual(
            summary["results"][0]["identity"]["operation_metrics_status"],
            "verified_equal",
        )

        malformed = copy.deepcopy(reports)
        malformed[0]["results"][0]["operation_metrics"]["source"][
            "logical_read_calls"
        ]["values"] = [1]
        with self.assertRaisesRegex(
            perf_abba_summary.AbbaSummaryInputError,
            "logical_read_calls.values has 1 samples",
        ):
            perf_abba_summary.summarize_reports(malformed)

        current = with_operation_metrics(four_legs())
        mixed = with_legacy_operation_metrics(four_legs())
        mixed[1]["results"][0]["operation_metrics"] = current[1]["results"][0][
            "operation_metrics"
        ]
        with self.assertRaisesRegex(
            perf_abba_summary.AbbaSummaryInputError,
            "operation_metrics identity",
        ):
            perf_abba_summary.summarize_reports(mixed)

    def test_filesystem_evidence_binds_complete_corpus_tool_and_configuration_identity(self):
        summary = perf_abba_summary.summarize_reports(with_filesystem_evidence(four_legs()))
        self.assertTrue(summary["verification"]["filesystem_evidence_identity_verified"])

        mutations = (
            (
                lambda legs: legs[1]["filesystem_evidence"][0]["corpus"].update(
                    name="changed-corpus"
                ),
                "case/corpus identity",
            ),
            (
                lambda legs: legs[1]["filesystem_evidence"][0]["tool"].update(
                    version="changed-tool"
                ),
                "tool identity",
            ),
            (
                lambda legs: legs[1]["filesystem_evidence"][0]["configuration"].update(
                    warmup_iterations_per_case=2
                ),
                "configuration identity",
            ),
        )
        for mutation, message in mutations:
            legs = with_filesystem_evidence(four_legs())
            mutation(legs)
            with self.subTest(message=message), self.assertRaisesRegex(
                perf_abba_summary.AbbaSummaryInputError, message
            ):
                perf_abba_summary.summarize_reports(legs)

        numeric_change = with_filesystem_evidence(four_legs())
        sample = numeric_change[1]["filesystem_evidence"][0]["samples"][0]
        sample["logical_read_calls"] = 999
        sample["logical_read_request_sizes"] = [1, 2, 3]
        self.assertTrue(
            perf_abba_summary.summarize_reports(numeric_change)["verification"][
                "filesystem_evidence_identity_verified"
            ]
        )

    def test_filesystem_range_size_pair_accepts_schema_one_legacy_reports(self):
        reports = with_filesystem_evidence(four_legs())
        for leg in reports:
            for sample in leg["filesystem_evidence"][0]["samples"]:
                sample.pop("logical_read_largest_requested_bytes")
                sample.pop("logical_read_largest_returned_bytes")

        summary = perf_abba_summary.summarize_reports(reports)
        self.assertTrue(summary["verification"]["filesystem_evidence_identity_verified"])

    def test_filesystem_range_size_pair_rejects_partial_or_mixed_legacy_shapes(self):
        partial = with_filesystem_evidence(four_legs())
        partial[0]["filesystem_evidence"][0]["samples"][0].pop(
            "logical_read_largest_returned_bytes"
        )
        with self.assertRaisesRegex(
            perf_abba_summary.AbbaSummaryInputError,
            "both logical read range-size counters or neither",
        ):
            perf_abba_summary.summarize_reports(partial)

        mixed = with_filesystem_evidence(four_legs())
        for sample in mixed[0]["filesystem_evidence"][0]["samples"]:
            sample.pop("logical_read_largest_requested_bytes")
            sample.pop("logical_read_largest_returned_bytes")
        with self.assertRaisesRegex(
            perf_abba_summary.AbbaSummaryInputError,
            "filesystem_evidence identity differs",
        ):
            perf_abba_summary.summarize_reports(mixed)

        within_evidence = with_filesystem_evidence(four_legs())
        for sample in within_evidence[0]["filesystem_evidence"][0]["samples"][1:]:
            sample.pop("logical_read_largest_requested_bytes")
            sample.pop("logical_read_largest_returned_bytes")
        with self.assertRaisesRegex(
            perf_abba_summary.AbbaSummaryInputError,
            "one logical read range-size counter schema consistently",
        ):
            perf_abba_summary.summarize_reports(within_evidence)

    def test_xlsx_repeated_store_primary_accepts_positive_control_and_zero_candidate_deltas(self):
        reports = with_xlsx_repeat_store_evidence(child_process_ids=True)
        summary = perf_abba_summary.summarize_reports(reports)
        self.assertEqual(summary["verification"]["result_count"], 1)
        self.assertEqual(summary["results"][0]["case"], "xlsx_source_repeated_store_medium")
        self.assertTrue(summary["verification"]["filesystem_evidence_identity_verified"])
        self.assertEqual(
            summary["results"][0]["elapsed_ns"]["accepted_statistics"],
            ["p50", "mean", "p95", "p99"],
        )

        timing_only = copy.deepcopy(reports)
        for sample in timing_only[1]["filesystem_evidence"][0]["samples"]:
            repeat = sample["xlsx_repeat_store"]
            repeat["query_elapsed_ns"] = [2, 3, 4, 5]
            repeat["timed_elapsed_total_ns"] = 14
            sample["elapsed_ns"] = 14
        timing_only[1]["results"][0]["elapsed_ns"] = elapsed([14] * 15)
        timing_only[1]["results"][0]["elapsed_ns"]["sample_order"] = list(range(15))
        self.assertEqual(
            perf_abba_summary.summarize_reports(timing_only)["verification"][
                "filesystem_evidence_identity_verified"
            ],
            True,
        )

        structural_constant = copy.deepcopy(reports)
        structural_constant[1]["filesystem_evidence"][0]["samples"][0][
            "xlsx_repeat_store"
        ]["cache_max_entries"] = 3
        with self.assertRaisesRegex(
            perf_abba_summary.AbbaSummaryInputError, "cache limits|identity differs"
        ):
            perf_abba_summary.summarize_reports(structural_constant)

    def test_xlsx_projection_separates_abba_role_from_diagnostic_location(self):
        reports = with_xlsx_repeat_store_evidence(child_process_ids=True)
        projections = {
            role: perf_abba_summary._project_report(
                report,
                f"package/abba-0269/{role}",
                profile=perf_abba_summary.detect_report_profile(report, role),
                report_role=role,
            )
            for role, report in zip(perf_abba_summary.LEG_ORDER, reports)
        }
        self.assertEqual(
            perf_abba_summary._summarize_projected_reports(projections),
            perf_abba_summary.summarize_reports(reports),
        )

        malformed = copy.deepcopy(reports[0])
        for sample in malformed["filesystem_evidence"][0]["samples"]:
            repeated_store = sample["xlsx_repeat_store"]
            repeated_store["diagnostics_delta"] = {
                field: 0 for field in repeated_store["diagnostics_delta"]
            }
            repeated_store["diagnostics_after"] = copy.deepcopy(
                repeated_store["diagnostics_before"]
            )
        with self.assertRaisesRegex(
            perf_abba_summary.AbbaSummaryInputError,
            r"package/abba-0269/a1.*positive control evidence",
        ):
            perf_abba_summary._project_report(
                malformed,
                "package/abba-0269/a1",
                profile=perf_abba_summary.detect_report_profile(malformed, "a1"),
                report_role="a1",
            )

    def test_xlsx_repeated_store_timing_binds_to_sample_order_and_result(self):
        reports = with_xlsx_repeat_store_evidence(child_process_ids=True)
        primary_sample = reports[0]["filesystem_evidence"][0]["samples"][0]
        primary_repeat = primary_sample["xlsx_repeat_store"]
        primary_repeat["timed_elapsed_total_ns"] += 1
        with self.assertRaisesRegex(
            perf_abba_summary.AbbaSummaryInputError,
            "does not equal query_elapsed_ns",
        ):
            perf_abba_summary.summarize_reports(reports)

        reports = with_xlsx_repeat_store_evidence(child_process_ids=True)
        primary_sample = reports[0]["filesystem_evidence"][0]["samples"][0]
        primary_sample["elapsed_ns"] += 1
        with self.assertRaisesRegex(
            perf_abba_summary.AbbaSummaryInputError,
            "must equal filesystem elapsed",
        ):
            perf_abba_summary.summarize_reports(reports)

        reports = with_xlsx_repeat_store_evidence(child_process_ids=True)
        elapsed_value = reports[0]["results"][0]["elapsed_ns"]
        elapsed_value["samples"][0] += 1
        with self.assertRaisesRegex(
            perf_abba_summary.AbbaSummaryInputError,
            "must equal authoritative result elapsed",
        ):
            perf_abba_summary.summarize_reports(reports)

        reports = with_xlsx_repeat_store_evidence(child_process_ids=True)
        reports[0]["results"][0]["elapsed_ns"]["sample_order"][1] = 0
        with self.assertRaisesRegex(
            perf_abba_summary.AbbaSummaryInputError,
            "sample_order must be an exact permutation",
        ):
            perf_abba_summary.summarize_reports(reports)

    def test_xlsx_repeated_store_requires_warm_fresh_samples_and_global_pids(self):
        mutations = (
            (
                lambda reports: reports[0]["filesystem_evidence"][0].update(
                    cache_states=["cold", "warm"]
                ),
                "cache_states must be exactly",
            ),
            (
                lambda reports: reports[0]["filesystem_evidence"][0].update(
                    fresh_child_per_sample=False
                ),
                "fresh_child_per_sample must be true",
            ),
        )
        for mutation, message in mutations:
            reports = with_xlsx_repeat_store_evidence(child_process_ids=True)
            mutation(reports)
            with self.subTest(message=message), self.assertRaisesRegex(
                perf_abba_summary.AbbaSummaryInputError, message
            ):
                perf_abba_summary.summarize_reports(reports)

        reports = with_xlsx_repeat_store_evidence(child_process_ids=True)
        duplicate_pid = reports[0]["filesystem_evidence"][0]["samples"][0][
            "child_process_id"
        ]
        reports[1]["filesystem_evidence"][0]["samples"][0][
            "child_process_id"
        ] = duplicate_pid
        with self.assertRaisesRegex(
            perf_abba_summary.AbbaSummaryInputError,
            "globally unique across ABBA legs",
        ):
            perf_abba_summary.summarize_reports(reports)

    def test_xlsx_repeated_store_rejects_semantic_source_string_and_schema_mutations(self):
        mutations = (
            (
                lambda reports: reports[1]["filesystem_evidence"][0]["samples"][0].update(
                    xlsx_semantic_sha256="e" * 64
                ),
                "semantic hash",
            ),
            (
                lambda reports: reports[1]["filesystem_evidence"][0]["samples"][0].update(
                    xlsx_source_sha256="e" * 64
                ),
                "source hash",
            ),
            (
                lambda reports: reports[1]["filesystem_evidence"][0]["samples"][0][
                    "xlsx_repeat_store"
                ]["query_names"].__setitem__(0, "changed"),
                "query_names",
            ),
            (
                lambda reports: reports[1]["filesystem_evidence"][0]["samples"][0][
                    "xlsx_repeat_store"
                ].update(extra=True),
                "keys mismatch",
            ),
        )
        for mutation, message in mutations:
            reports = with_xlsx_repeat_store_evidence(child_process_ids=True)
            mutation(reports)
            with self.subTest(message=message), self.assertRaisesRegex(
                perf_abba_summary.AbbaSummaryInputError, message
            ):
                perf_abba_summary.summarize_reports(reports)

    def test_xlsx_repeated_store_rejects_stale_missing_extra_wrong_type_and_counter_arithmetic(self):
        mutations = (
            (
                lambda reports: reports[0]["filesystem_evidence"][0]["samples"][0][
                    "xlsx_repeat_store"
                ].update(timed_elapsed_total_ns=11),
                "does not equal query_elapsed_ns",
            ),
            (
                lambda reports: reports[0]["filesystem_evidence"][0]["samples"][0][
                    "xlsx_repeat_store"
                ]["diagnostics_after"].update(cache_evictions=999),
                "inconsistent for cache_evictions",
            ),
            (
                lambda reports: reports[0]["filesystem_evidence"][0]["samples"][0][
                    "xlsx_repeat_store"
                ].pop("diagnostics_delta"),
                "keys mismatch",
            ),
            (
                lambda reports: reports[0]["filesystem_evidence"][0]["samples"][0][
                    "xlsx_repeat_store"
                ].update(query_iterations="8"),
                "must be an unsigned 64-bit integer",
            ),
        )
        for mutation, message in mutations:
            reports = with_xlsx_repeat_store_evidence()
            mutation(reports)
            with self.subTest(message=message), self.assertRaisesRegex(
                perf_abba_summary.AbbaSummaryInputError, message
            ):
                perf_abba_summary.summarize_reports(reports)

    def test_xlsx_repeated_store_allocation_and_child_pid_shapes_are_strict(self):
        reports = with_xlsx_repeat_store_evidence(child_process_ids=True)
        for leg in reports:
            for sample in leg["filesystem_evidence"][0]["samples"]:
                sample["allocation_metrics"] = {
                    "status": "unavailable",
                    "scope": "operation_global_system_allocator",
                }
        self.assertEqual(
            perf_abba_summary.summarize_reports(reports)["verification"]["result_count"],
            1,
        )

        malformed = copy.deepcopy(reports)
        malformed[1]["filesystem_evidence"][0]["samples"][0][
            "allocation_metrics"
        ] = {"status": "unavailable", "scope": "wrong"}
        with self.assertRaisesRegex(
            perf_abba_summary.AbbaSummaryInputError, "allocator schema|scope"
        ):
            perf_abba_summary.summarize_reports(malformed)

        malformed = with_xlsx_repeat_store_evidence(child_process_ids=True)
        malformed[0]["filesystem_evidence"][0]["samples"][0]["child_process_id"] = 0
        with self.assertRaisesRegex(
            perf_abba_summary.AbbaSummaryInputError, "child_process_id.*positive"
        ):
            perf_abba_summary.summarize_reports(malformed)

    def test_xlsx_repeated_store_structural_controls_are_excluded_and_cannot_masquerade(self):
        primary = with_xlsx_repeat_store_evidence()
        structural = with_xlsx_repeat_store_evidence(structural=True, pid_offset=10_000)
        for primary_leg, structural_leg in zip(primary, structural):
            primary_leg["results"].append(structural_leg["results"][0])
            primary_leg["filesystem_evidence"].append(
                structural_leg["filesystem_evidence"][0]
            )
            primary_leg["configuration"]["cases"].append(
                structural_leg["configuration"]["cases"][0]
            )
            shape = structural_leg["configuration"]["corpus_shapes"][0]
            if shape not in primary_leg["configuration"]["corpus_shapes"]:
                primary_leg["configuration"]["corpus_shapes"].append(shape)
            for evidence in primary_leg["filesystem_evidence"]:
                evidence["configuration"] = copy.deepcopy(primary_leg["configuration"])
            # Structural rows are validated for their own evidence contract but
            # never enter the primary elapsed/source/sink summary path.
            primary_leg["results"][1]["elapsed_ns"]["samples"] = [999] * 15
            primary_leg["results"][1]["source"] = {"arbitrary": "structural-only"}
        summary = perf_abba_summary.summarize_reports(primary)
        self.assertEqual(summary["verification"]["result_count"], 1)
        self.assertEqual(
            [result["case"] for result in summary["results"]],
            ["xlsx_source_repeated_store_medium"],
        )

        masquerading = with_xlsx_repeat_store_evidence(structural=True)
        masquerading[0]["filesystem_evidence"][0]["samples"][0][
            "xlsx_repeat_store"
        ]["implementation"] = "source_backed_cached_store"
        with self.assertRaisesRegex(
            perf_abba_summary.AbbaSummaryInputError, "does not match structural"
        ):
            perf_abba_summary.summarize_reports(masquerading)

    def test_xlsx_repeated_store_selector_and_corpus_contracts_reject_renames_or_arbitrary_corpus(self):
        mutations = (
            (
                lambda reports: (
                    reports[0]["results"][0].update(
                        case="xlsx_source_repeated_store_medium_renamed"
                    ),
                    reports[0]["filesystem_evidence"][0].update(
                        case="xlsx_source_repeated_store_medium_renamed"
                    ),
                ),
                "pinned repeated-store corpus|not permitted on filesystem case",
            ),
            (
                lambda reports: (
                    reports[0]["results"][0].update(
                        case="xlsx_source_repeated_store_medium"
                    ),
                    reports[0]["filesystem_evidence"][0].update(
                        case="xlsx_source_repeated_store_medium"
                    ),
                    reports[0]["filesystem_evidence"][0]["samples"][0][
                        "xlsx_repeat_store"
                    ].update(
                        implementation="explicit_part_data_reacquisition_structural_control"
                    ),
                ),
                "does not match",
            ),
            (
                lambda reports: (
                    reports[1]["results"][0]["corpus"].update(name="arbitrary-corpus"),
                    reports[1]["filesystem_evidence"][0]["corpus"].update(
                        name="arbitrary-corpus"
                    ),
                ),
                "pinned.*corpus",
            ),
            (
                lambda reports: (
                    reports[1]["results"][0]["corpus"].update(archive_bytes=1),
                    reports[1]["filesystem_evidence"][0]["corpus"].update(
                        archive_bytes=1
                    ),
                ),
                "pinned.*corpus",
            ),
            (
                lambda reports: (
                    reports[1]["results"][0]["corpus"].update(
                        archive_sha256="e" * 64
                    ),
                    reports[1]["filesystem_evidence"][0]["corpus"].update(
                        archive_sha256="e" * 64
                    ),
                ),
                "pinned.*corpus",
            ),
        )
        for mutation, message in mutations:
            reports = with_xlsx_repeat_store_evidence()
            mutation(reports)
            with self.subTest(message=message), self.assertRaisesRegex(
                perf_abba_summary.AbbaSummaryInputError, message
            ):
                perf_abba_summary.summarize_reports(reports)

    def test_xlsx_repeated_store_pins_full_corpus_and_distinct_semantic_identities(self):
        smoke_path = Path("/tmp/xlsx-reconcile-A-four.aVHKkD.json")
        if smoke_path.exists():
            smoke = json.loads(smoke_path.read_text())
            for evidence in smoke["filesystem_evidence"]:
                case = evidence["case"]
                self.assertEqual(
                    evidence["corpus"],
                    perf_abba_summary.FIXED_CASE_CORPUS_IDENTITIES[case],
                )
                sample = evidence["samples"][0]
                self.assertEqual(
                    sample["xlsx_semantic_sha256"],
                    "020fdd140d2959ea4f480676a3d4d0bf840927e25251cb6cad37a043ab80627e",
                )
                self.assertEqual(
                    sample["xlsx_repeat_store"]["semantic_projection_sha256"],
                    "01c253bf3fc611835e0806414c6417a9cfbb012ff6e01f9bb55cec94236a6235",
                )

        reports = with_xlsx_repeat_store_evidence(child_process_ids=True)
        sample = reports[0]["filesystem_evidence"][0]["samples"][0]
        self.assertNotEqual(
            sample["xlsx_semantic_sha256"],
            sample["xlsx_repeat_store"]["semantic_projection_sha256"],
        )
        nested_mutations = (
            lambda corpus: corpus["xlsx"]["source_members"]["worksheets"].append(
                "xl/worksheets/extra.xml"
            ),
            lambda corpus: corpus["xlsx"].update(extra_nested=True),
            lambda corpus: corpus.update(target_payload_bytes=2),
        )
        for mutate in nested_mutations:
            mutated = with_xlsx_repeat_store_evidence(child_process_ids=True)
            corpus = mutated[1]["filesystem_evidence"][0]["corpus"]
            mutate(corpus)
            mutated[1]["results"][0]["corpus"] = copy.deepcopy(corpus)
            with self.assertRaisesRegex(
                perf_abba_summary.AbbaSummaryInputError,
                "does not match the pinned|pinned.*corpus",
            ):
                perf_abba_summary.summarize_reports(mutated)

    def test_xlsx_repeated_store_corpus_rename_cannot_downgrade_to_ordinary_case(self):
        reports = with_xlsx_repeat_store_evidence(child_process_ids=True)
        renamed_case = "ordinary_case_with_repeated_corpus"
        reports[0]["results"][0]["case"] = renamed_case
        reports[0]["filesystem_evidence"][0]["case"] = renamed_case
        reports[0]["configuration"]["cases"] = [renamed_case]
        with self.assertRaisesRegex(
            perf_abba_summary.AbbaSummaryInputError,
            "pinned repeated-store corpus",
        ):
            perf_abba_summary.summarize_reports(reports)

    def test_xlsx_repeated_store_marker_does_not_claim_shared_ordinary_archive(self):
        # The medium repeated-store archive is byte-for-byte shared with the
        # ordinary xlsx-cell-values-medium claim.  Archive SHA alone therefore
        # cannot dispatch legacy evidence into the repeated-store contract.
        ordinary_corpus = {
            "name": "xlsx-cell-values-medium",
            "generator": "litchi-xlsx-cell-values-source-edit-media-multi-sheet-v1",
            "archive_sha256": (
                "dfff7ec0c749d9e404091776f15a8fb690985af7f58efdfe659dbeaed7145036"
            ),
        }
        self.assertFalse(
            perf_abba_summary._looks_like_xlsx_repeat_store_corpus(ordinary_corpus)
        )

    def test_xlsx_repeated_store_full_rewrite_is_a_different_selector_and_hash(self):
        """Generic JSON validation cannot authenticate a self-consistent rewrite.

        The requested selector and the raw report canonical hashes are the
        protocol boundary: a full rewrite that removes every repeated-store
        marker may be summarized as a new generic claim, but it cannot satisfy
        a package/registry claim that requested the original selector and pins
        the original raw hashes.
        """

        reports = with_xlsx_repeat_store_evidence(child_process_ids=True)
        original_hashes = [
            perf_abba_summary._canonical_sha256(report, f"original[{index}]")
            for index, report in enumerate(reports)
        ]
        generic_corpus = {
            "name": "ordinary-rewritten-medium",
            "generator": "ordinary-generator-v1",
            "shape": "medium",
            "archive_sha256": "f" * 64,
        }
        rewritten_case = "ordinary_rewritten_case"
        for report in reports:
            result = report["results"][0]
            result["case"] = rewritten_case
            result["corpus"] = copy.deepcopy(generic_corpus)
            report["configuration"]["cases"] = [rewritten_case]
            report["configuration"]["corpus_shapes"] = ["medium"]
            evidence = report["filesystem_evidence"][0]
            evidence["case"] = rewritten_case
            evidence["corpus"] = copy.deepcopy(generic_corpus)
            evidence["configuration"] = copy.deepcopy(report["configuration"])
            for sample in evidence["samples"]:
                for key in (
                    "xlsx_repeat_store",
                    "xlsx_source_sha256",
                    "xlsx_semantic_sha256",
                    "child_process_id",
                ):
                    sample.pop(key, None)

        rewritten_hashes = [
            perf_abba_summary._canonical_sha256(report, f"rewritten[{index}]")
            for index, report in enumerate(reports)
        ]
        self.assertTrue(
            all(
                original != rewritten
                for original, rewritten in zip(original_hashes, rewritten_hashes)
            )
        )
        self.assertEqual(
            perf_abba_summary.summarize_reports(reports)["verification"]["result_count"],
            1,
        )
        with self.assertRaisesRegex(
            perf_abba_summary.AbbaSummaryInputError,
            "selectors did not match any case/corpus result",
        ):
            perf_abba_summary.summarize_reports(
                reports, cases=["xlsx_source_repeated_store_medium"]
            )

    def test_xlsx_repeated_store_rejects_forged_primary_result_channels(self):
        for field, value in (
            ("source", {"forged": True}),
            ("sink", {"forged": True}),
            ("output_sha256", "e" * 64),
        ):
            reports = with_xlsx_repeat_store_evidence(child_process_ids=True)
            reports[0]["results"][0][field] = value
            with self.subTest(field=field), self.assertRaisesRegex(
                perf_abba_summary.AbbaSummaryInputError,
                rf"results\[xlsx_source_repeated_store_medium\].{field} must be absent or null",
            ):
                perf_abba_summary.summarize_reports(reports)

    def test_xlsx_repeated_store_per_sample_scenario_size_and_pid_contracts_fail_closed(self):
        mutations = (
            (
                lambda reports: reports[1]["filesystem_evidence"][0]["samples"][1][
                    "xlsx_repeat_store"
                ].update(scenario="oversized"),
                "scenario does not match",
            ),
            (
                lambda reports: reports[1]["filesystem_evidence"][0]["samples"][1][
                    "xlsx_repeat_store"
                ].update(selected_member_uncompressed_bytes=63_295),
                "selected_member_uncompressed_bytes does not match",
            ),
            (
                lambda reports: reports[1]["filesystem_evidence"][0]["samples"][1].pop(
                    "child_process_id"
                ),
                "must contain child_process_id",
            ),
            (
                lambda reports: reports[1]["filesystem_evidence"][0]["samples"][1].update(
                    child_process_id=reports[1]["filesystem_evidence"][0]["samples"][0][
                        "child_process_id"
                    ]
                ),
                "child_process_id values must be unique",
            ),
        )
        for mutation, message in mutations:
            reports = with_xlsx_repeat_store_evidence()
            mutation(reports)
            with self.subTest(message=message), self.assertRaisesRegex(
                perf_abba_summary.AbbaSummaryInputError, message
            ):
                perf_abba_summary.summarize_reports(reports)

        missing = with_xlsx_repeat_store_evidence()
        for sample in missing[0]["filesystem_evidence"][0]["samples"]:
            sample.pop("xlsx_repeat_store")
        with self.assertRaisesRegex(
            perf_abba_summary.AbbaSummaryInputError, "must contain xlsx_repeat_store"
        ):
            perf_abba_summary.summarize_reports(missing)

    def test_default_drift_ceilings_and_custom_ceilings_are_applied_per_statistic(self):
        legs = reports_for_values(
            (
                [100, 100, 100, 100, 100],
                [80, 80, 80, 80, 80],
                [80, 80, 80, 80, 80],
                [106, 106, 106, 106, 106],
            )
        )
        default = perf_abba_summary.summarize_reports(legs)["results"][0]["elapsed_ns"]
        self.assertEqual(default["accepted_statistics"], ["p95", "p99"])
        self.assertIn("p50", default["rejected_statistics"])
        custom = perf_abba_summary.summarize_reports(
            legs,
            drift_ceilings={"p50": 10, "mean": 10, "p95": 10, "p99": 15},
        )["results"][0]["elapsed_ns"]
        self.assertEqual(custom["accepted_statistics"], ["p50", "mean", "p95", "p99"])

    def test_adverse_both_and_sign_disagreement_are_classified(self):
        adverse = reports_for_values(
            (
                [100, 100, 100, 100, 100],
                [120, 120, 120, 120, 120],
                [130, 130, 130, 130, 130],
                [110, 110, 110, 110, 110],
            )
        )
        elapsed_summary = perf_abba_summary.summarize_reports(adverse)["results"][0][
            "elapsed_ns"
        ]
        self.assertEqual(elapsed_summary["adverse_both_statistics"], [
            "p50",
            "mean",
            "p95",
            "p99",
        ])
        self.assertTrue(
            all(
                reason.startswith("candidate is not lower in both paired directions")
                for reason in elapsed_summary["rejected_statistics"].values()
            )
        )

        mixed = reports_for_values(
            (
                [100, 100, 100, 100, 100],
                [80, 80, 80, 80, 80],
                [120, 120, 120, 120, 120],
                [100, 100, 100, 100, 100],
            )
        )
        elapsed_summary = perf_abba_summary.summarize_reports(mixed)["results"][0][
            "elapsed_ns"
        ]
        self.assertEqual(elapsed_summary["adverse_both_statistics"], [])
        self.assertTrue(
            all(
                reason.startswith("paired directions disagree")
                for reason in elapsed_summary["rejected_statistics"].values()
            )
        )

        tie_and_adverse = reports_for_values(
            (
                [100, 100, 100, 100, 100],
                [100, 100, 100, 100, 100],
                [120, 120, 120, 120, 120],
                [100, 100, 100, 100, 100],
            )
        )
        elapsed_summary = perf_abba_summary.summarize_reports(tie_and_adverse)[
            "results"
        ][0]["elapsed_ns"]
        self.assertTrue(
            all(
                reason.startswith("candidate is not lower in both paired directions")
                for reason in elapsed_summary["rejected_statistics"].values()
            )
        )

    def test_environment_provenance_allows_expected_variants_and_rejects_stable_drift(self):
        legs = four_legs()
        for leg, revision in zip(
            legs,
            ("control-revision-2", "candidate-revision-2", "candidate-revision-2", "control-revision-2"),
        ):
            leg["environment"]["git_revision"] = revision
        summary = perf_abba_summary.summarize_reports(legs)
        self.assertEqual(
            [
                summary["environment"]["legs"][label]["git_revision"]
                for label in ("a1", "b1", "b2", "a2")
            ],
            [
                "control-revision-2",
                "candidate-revision-2",
                "candidate-revision-2",
                "control-revision-2",
            ],
        )
        self.assertEqual(summary["verification"]["environment_legs_recorded"], True)

        legs = four_legs()
        legs[1]["environment"]["allocator"] = "different allocator"
        with self.assertRaisesRegex(
            perf_abba_summary.AbbaSummaryInputError, "stable environment identity"
        ):
            perf_abba_summary.summarize_reports(legs)

    def test_paired_implementation_identity_requires_clean_distinct_revisions(self):
        mutations = (
            (lambda legs: legs[3]["environment"].update(git_revision="other-control"),
             "control A1/A2 git_revision"),
            (lambda legs: legs[2]["environment"].update(git_revision="other-candidate"),
             "candidate B1/B2 git_revision"),
            (
                lambda legs: (
                    legs[1]["environment"].update(git_revision="control-revision"),
                    legs[2]["environment"].update(git_revision="control-revision"),
                ),
                "distinct",
            ),
            (lambda legs: legs[0]["environment"].update(git_revision=""), "git_revision"),
            (lambda legs: legs[0]["environment"].update(git_worktree_dirty=True),
             "git_worktree_dirty"),
        )
        for mutation, message in mutations:
            legs = four_legs()
            mutation(legs)
            with self.subTest(message=message), self.assertRaisesRegex(
                perf_abba_summary.AbbaSummaryInputError, message
            ):
                perf_abba_summary.summarize_reports(legs)

        summary = perf_abba_summary.summarize_reports(four_legs())
        identity = summary["implementation_identity"]
        self.assertEqual(identity["control"]["git_revision"], "control-revision")
        self.assertEqual(identity["candidate"]["git_revision"], "candidate-revision")
        self.assertEqual(identity["control"]["legs"], ["a1", "a2"])
        self.assertEqual(identity["candidate"]["legs"], ["b1", "b2"])
        self.assertEqual(identity["control"]["binary_sha256"], "a" * 64)
        self.assertEqual(identity["candidate"]["binary_sha256"], "b" * 64)
        self.assertTrue(identity["distinct"])

    def test_binary_identity_is_required_exact_within_legs_and_distinct_across_legs(self):
        missing = four_legs()
        missing[0].pop("binary_identity")
        with self.assertRaisesRegex(
            perf_abba_summary.AbbaSummaryInputError, "binary_identity"
        ):
            perf_abba_summary.summarize_reports(missing)

        malformed = four_legs()
        malformed[1]["binary_identity"]["binary_sha256"] = "not-a-sha256"
        with self.assertRaisesRegex(
            perf_abba_summary.AbbaSummaryInputError, "binary_sha256"
        ):
            perf_abba_summary.summarize_reports(malformed)

        non_executable_mode = four_legs()
        non_executable_mode[0]["binary_identity"]["mode_bits"] = 0o644
        with self.assertRaisesRegex(
            perf_abba_summary.AbbaSummaryInputError, "executable permission bits"
        ):
            perf_abba_summary.summarize_reports(non_executable_mode)

        missing_unix_mode = four_legs()
        missing_unix_mode[0]["binary_identity"]["mode_bits"] = None
        with self.assertRaisesRegex(
            perf_abba_summary.AbbaSummaryInputError, "present for Unix targets"
        ):
            perf_abba_summary.summarize_reports(missing_unix_mode)

        oversized = four_legs()
        oversized[0]["binary_identity"]["binary_bytes"] = 1 << 64
        with self.assertRaisesRegex(
            perf_abba_summary.AbbaSummaryInputError, "positive unsigned integer"
        ):
            perf_abba_summary.summarize_reports(oversized)

        same_leg_drift = four_legs()
        same_leg_drift[3]["binary_identity"]["binary_bytes"] += 1
        with self.assertRaisesRegex(
            perf_abba_summary.AbbaSummaryInputError, "control binary identity"
        ):
            perf_abba_summary.summarize_reports(same_leg_drift)

        identical = four_legs()
        identical[1]["binary_identity"] = copy.deepcopy(
            identical[0]["binary_identity"]
        )
        identical[2]["binary_identity"] = copy.deepcopy(
            identical[0]["binary_identity"]
        )
        with self.assertRaisesRegex(
            perf_abba_summary.AbbaSummaryInputError, "hashes must differ"
        ):
            perf_abba_summary.summarize_reports(identical)

        summary = perf_abba_summary.summarize_reports(four_legs())
        self.assertTrue(summary["verification"]["binary_identity_verified"])
        self.assertTrue(summary["verification"]["binary_hashes_distinct"])

    def test_rust_statistics_verify_integer_tails_dispersion_and_uncertainty(self):
        values = list(range(1, 17))
        expected = elapsed(values)
        recomputed = perf_abba_summary.recompute_statistics(expected, "test.elapsed_ns")
        self.assertEqual(recomputed["sample_count"], 16)
        self.assertEqual(recomputed["min"], 1)
        self.assertEqual(recomputed["p50"], 8)
        self.assertEqual(recomputed["p95"], 16)
        self.assertEqual(recomputed["p99"], 16)
        self.assertEqual(recomputed["max"], 16)
        self.assertAlmostEqual(recomputed["mean"], expected["mean"])
        self.assertAlmostEqual(
            recomputed["standard_deviation"], expected["standard_deviation"]
        )
        self.assertAlmostEqual(
            recomputed["confidence_interval_95"]["lower"],
            expected["confidence_interval_95"]["lower"],
        )
        self.assertAlmostEqual(
            recomputed["confidence_interval_95"]["upper"],
            expected["confidence_interval_95"]["upper"],
        )
        self.assertEqual(
            recomputed["confidence_interval_95"]["method"],
            "two-sided Student's t interval for the mean",
        )

        for field in ("min", "p50", "p95", "p99", "max"):
            malformed = copy.deepcopy(expected)
            malformed[field] += 1
            with self.subTest(field=field), self.assertRaisesRegex(
                perf_abba_summary.AbbaSummaryInputError, "disagrees"
            ):
                perf_abba_summary.recompute_statistics(malformed, "test.elapsed_ns")
        for field in ("mean", "standard_deviation"):
            malformed = copy.deepcopy(expected)
            malformed[field] += 1.0
            with self.subTest(field=field), self.assertRaisesRegex(
                perf_abba_summary.AbbaSummaryInputError, "disagrees"
            ):
                perf_abba_summary.recompute_statistics(malformed, "test.elapsed_ns")
        malformed = copy.deepcopy(expected)
        malformed["confidence_interval_95"]["lower"] += 1.0
        with self.assertRaisesRegex(perf_abba_summary.AbbaSummaryInputError, "disagrees"):
            perf_abba_summary.recompute_statistics(malformed, "test.elapsed_ns")
        malformed = copy.deepcopy(expected)
        malformed["confidence_interval_95"]["method"] = "normal interval"
        with self.assertRaisesRegex(perf_abba_summary.AbbaSummaryInputError, "harness"):
            perf_abba_summary.recompute_statistics(malformed, "test.elapsed_ns")

    def test_typed_harness_fields_warmups_samples_and_shapes_fail_closed(self):
        mutations = (
            (lambda legs: legs[0]["tool"].update(name="other-tool"), "tool.name"),
            (lambda legs: legs[0]["tool"].update(profile=1), "tool.profile"),
            (lambda legs: legs[0]["environment"].update(logical_cpus_available=True),
             "logical_cpus_available"),
            (lambda legs: legs[0]["environment"].pop("allocator"), "allocator"),
            (lambda legs: legs[0]["configuration"].update(warmup_iterations_per_case=0),
             "warmup_iterations_per_case"),
            (lambda legs: legs[0]["configuration"].update(samples_per_case=14),
             "at least 15"),
            (lambda legs: legs[0]["configuration"].update(cases="synthetic_case"),
             "configuration.cases"),
            (lambda legs: legs[0]["configuration"].update(corpus_shapes=["tiny"]),
             "configuration"),
            (lambda legs: legs[0]["results"][0]["corpus"].update(shape="unknown"),
             "shape declarations"),
        )
        for mutation, message in mutations:
            legs = four_legs()
            mutation(legs)
            with self.subTest(message=message), self.assertRaisesRegex(
                perf_abba_summary.AbbaSummaryInputError, message
            ):
                perf_abba_summary.summarize_reports(legs)

    def test_xlsx_page_break_fixed_corpus_identity_is_exact(self):
        fixed_page_break = four_legs()
        for leg in fixed_page_break:
            leg["configuration"]["cases"] = [
                "xlsx_eager_page_break_edit_save",
                "xlsx_source_backed_page_break_edit_save",
            ]
            for index, result in enumerate(leg["results"]):
                result["case"] = leg["configuration"]["cases"][index]
                result["corpus"].update(
                    name="xlsx-page-break-media",
                    generator="litchi-xlsx-page-break-source-edit-media-v1",
                    shape="media-rich",
                )
        perf_abba_summary.summarize_reports(fixed_page_break)

        for leg in fixed_page_break:
            leg["results"][0]["corpus"]["generator"] = "unknown-generator"
        with self.assertRaisesRegex(
            perf_abba_summary.AbbaSummaryInputError, "cover result shapes"
        ):
            perf_abba_summary.summarize_reports(fixed_page_break)

    def test_opc_source_overlay_multi_part_fixed_shape_count_matrix_is_exact(self):
        cases = sorted(perf_abba_summary.OPC_SOURCE_OVERLAY_MULTI_PART_CASES)
        identities = perf_abba_summary.OPC_SOURCE_OVERLAY_MULTI_PART_CORPUS_IDENTITIES
        counts = perf_abba_summary.OPC_SOURCE_OVERLAY_MULTI_PART_COUNTS
        reports = four_legs()
        def populate(legs):
            for leg_index, leg in enumerate(legs):
                leg["configuration"]["cases"] = cases
                rows = []
                for case in cases:
                    mode = case.removeprefix("opc_source_overlay_multi_part_")
                    for identity in identities.values():
                        for count in counts:
                            result = copy.deepcopy(leg["results"][0])
                            result["case"] = case
                            result["corpus"] = {
                                field: value
                                for field, value in identity.items()
                                if field != "name_prefix"
                            }
                            result["corpus"].update(
                                name=f"{identity['name_prefix']}{count}",
                            )
                            source, output_digest = opc_multi_source(
                                result["elapsed_ns"]["samples"],
                                mode,
                                count,
                                source_shape=identity["shape"],
                                payload_kind=identity["payload_kind"],
                                uncompressed_payload_bytes=identity[
                                    "uncompressed_payload_bytes"
                                ],
                                archive_bytes=identity["archive_bytes"],
                                archive_sha256=identity["archive_sha256"],
                                expected_eager_sha256=(
                                    perf_abba_summary.OPC_SOURCE_OVERLAY_MULTI_PART_EXPECTED_EAGER_SHA256[
                                        (case, identity["shape"], count)
                                    ]
                                ),
                                offset=leg_index,
                            )
                            result["source"] = source
                            result["sink"] = opc_multi_sink()
                            result["elapsed_ns"]["sample_order"] = list(
                                range(len(result["elapsed_ns"]["samples"]))
                            )
                            result["output_sha256"] = output_digest
                            rows.append(result)
                leg["results"] = rows

        populate(reports)
        self.assertFalse(perf_abba_summary.OPC_SOURCE_OVERLAY_MULTI_PART_HASHES_PENDING)
        self.assertEqual(
            len(perf_abba_summary.OPC_SOURCE_OVERLAY_MULTI_PART_EXPECTED_EAGER_SHA256),
            27,
        )
        for identity in identities.values():
            for count in counts:
                corpus = {
                    field: value
                    for field, value in identity.items()
                    if field != "name_prefix"
                }
                corpus["name"] = f"{identity['name_prefix']}{count}"
                self.assertEqual(
                    perf_abba_summary._opc_source_overlay_multi_part_identity(corpus),
                    (identity["shape"], count),
                )
        summary = perf_abba_summary.summarize_reports(reports)
        self.assertEqual(summary["verification"]["result_count"], 27)

        missing_sink = copy.deepcopy(reports)
        missing_sink[0]["results"][0].pop("sink")
        with self.assertRaisesRegex(
            perf_abba_summary.AbbaSummaryInputError,
            "sink must be an object",
        ):
            perf_abba_summary.summarize_reports(missing_sink)

        malformed = copy.deepcopy(reports)
        malformed[0]["results"].pop()
        with self.assertRaisesRegex(
            perf_abba_summary.AbbaSummaryInputError,
            "complete multi-Part OPC 3-shape x 3-count matrix",
        ):
            perf_abba_summary.summarize_reports(malformed)

        arbitrary_hash = copy.deepcopy(reports)
        arbitrary_hash[0]["results"][0]["corpus"]["archive_sha256"] = "d" * 64
        with self.assertRaisesRegex(
            perf_abba_summary.AbbaSummaryInputError,
            "pinned multi-Part OPC identity|fixed identity",
        ):
            perf_abba_summary.summarize_reports(arbitrary_hash)

    def test_allocator_instrumentation_is_not_accepted_for_latency_abba(self):
        legs = four_legs()
        for leg in legs:
            leg["tool"]["instrumentation"] = "system_allocator"
        with self.assertRaisesRegex(
            perf_abba_summary.AbbaSummaryInputError,
            "instrumentation.*latency ABBA",
        ):
            perf_abba_summary.summarize_reports(legs)

    def test_opc_source_overlay_nested_schema_projection_and_sample_binding(self):
        identity = copy.deepcopy(
            perf_abba_summary.OPC_SOURCE_OVERLAY_MULTI_PART_CORPUS_IDENTITIES[
                "overlay-small"
            ]
        )
        corpus = {
            field: value for field, value in identity.items() if field != "name_prefix"
        }
        corpus.update(
            name=f"{identity['name_prefix']}2",
        )
        source, output_digest = opc_multi_source(
            [10, 12],
            "mixed",
            2,
            source_shape="overlay-small",
            payload_kind="compressible",
            uncompressed_payload_bytes=32 * 1024,
            archive_bytes=7451,
            archive_sha256=(
                "4338dea03f37b0ea2ad63a055fb5cfb7df79a5b0de864365e981e453e1a65509"
            ),
            expected_eager_sha256=(
                perf_abba_summary.OPC_SOURCE_OVERLAY_MULTI_PART_EXPECTED_EAGER_SHA256[
                    (
                        "opc_source_overlay_multi_part_mixed",
                        "overlay-small",
                        2,
                    )
                ]
            ),
            offset=5,
        )
        overlay = source["opc_source_overlay"]
        sink = opc_multi_sink()
        perf_abba_summary._validate_opc_source_overlay(
            overlay,
            "test.source.opc_source_overlay",
            case="opc_source_overlay_multi_part_mixed",
            corpus=corpus,
            samples_per_case=2,
            elapsed_samples=[10, 12],
            sample_order=[0, 1],
            source=source,
            sink=sink,
            output_sha256=output_digest,
            allow_pending_hashes=True,
        )
        projected = perf_abba_summary._source_identity_projection(source)
        self.assertNotIn("read_calls", projected)
        self.assertNotIn("preparation_ns", projected["opc_source_overlay"])
        self.assertEqual(
            projected["opc_source_overlay"]["implementation"],
            overlay["implementation"],
        )

        missing = copy.deepcopy(overlay)
        missing.pop("publication_ns")
        with self.assertRaisesRegex(
            perf_abba_summary.AbbaSummaryInputError, "schema mismatch"
        ):
            perf_abba_summary._validate_opc_source_overlay(
                missing,
                "test.source.opc_source_overlay",
                case="opc_source_overlay_multi_part_mixed",
                corpus=corpus,
                samples_per_case=2,
                elapsed_samples=[10, 12],
                sample_order=[0, 1],
                source=source,
                sink=sink,
                output_sha256=output_digest,
                allow_pending_hashes=True,
            )

        wrong_cardinality = copy.deepcopy(overlay)
        wrong_cardinality["publication_ns"] = [8]
        with self.assertRaisesRegex(
            perf_abba_summary.AbbaSummaryInputError, "exactly 2 samples"
        ):
            perf_abba_summary._validate_opc_source_overlay(
                wrong_cardinality,
                "test.source.opc_source_overlay",
                case="opc_source_overlay_multi_part_mixed",
                corpus=corpus,
                samples_per_case=2,
                elapsed_samples=[10, 12],
                sample_order=[0, 1],
                source=source,
                sink=sink,
                output_sha256=output_digest,
                allow_pending_hashes=True,
            )

        wrong_sum = copy.deepcopy(overlay)
        wrong_sum["publication_ns"][1] += 1
        with self.assertRaisesRegex(
            perf_abba_summary.AbbaSummaryInputError, "bind to sorted elapsed_ns"
        ):
            perf_abba_summary._validate_opc_source_overlay(
                wrong_sum,
                "test.source.opc_source_overlay",
                case="opc_source_overlay_multi_part_mixed",
                corpus=corpus,
                samples_per_case=2,
                elapsed_samples=[10, 12],
                sample_order=[0, 1],
                source=source,
                sink=sink,
                output_sha256=output_digest,
                allow_pending_hashes=True,
            )

        bad_source = copy.deepcopy(source)
        bad_source["read_calls"][0] += 1
        with self.assertRaisesRegex(
            perf_abba_summary.AbbaSummaryInputError, "disagrees with source.read_calls"
        ):
            perf_abba_summary._validate_opc_source_overlay(
                overlay,
                "test.source.opc_source_overlay",
                case="opc_source_overlay_multi_part_mixed",
                corpus=corpus,
                samples_per_case=2,
                elapsed_samples=[10, 12],
                sample_order=[0, 1],
                source=bad_source,
                sink=sink,
                output_sha256=output_digest,
                allow_pending_hashes=True,
            )

        bad_sink = copy.deepcopy(sink)
        bad_sink["largest_write"] = 65_537
        with self.assertRaisesRegex(
            perf_abba_summary.AbbaSummaryInputError, "largest_write exceeds"
        ):
            perf_abba_summary._validate_opc_source_overlay(
                overlay,
                "test.source.opc_source_overlay",
                case="opc_source_overlay_multi_part_mixed",
                corpus=corpus,
                samples_per_case=2,
                elapsed_samples=[10, 12],
                sample_order=[0, 1],
                source=source,
                sink=bad_sink,
                output_sha256=output_digest,
                allow_pending_hashes=True,
            )

        bad_eager = copy.deepcopy(overlay)
        bad_eager["expected_eager_sha256"] = "f" * 64
        with self.assertRaisesRegex(
            perf_abba_summary.AbbaSummaryInputError,
            "pinned case/shape/count identity",
        ):
            perf_abba_summary._validate_opc_source_overlay(
                bad_eager,
                "test.source.opc_source_overlay",
                case="opc_source_overlay_multi_part_mixed",
                corpus=corpus,
                samples_per_case=2,
                elapsed_samples=[10, 12],
                sample_order=[0, 1],
                source=source,
                sink=sink,
                output_sha256=output_digest,
                allow_pending_hashes=True,
            )

        bad_sink_formula = copy.deepcopy(overlay)
        bad_sink_formula["sink_max_bytes"] += 1
        with self.assertRaisesRegex(
            perf_abba_summary.AbbaSummaryInputError,
            r"2\*accepted_bytes\+65536",
        ):
            perf_abba_summary._validate_opc_source_overlay(
                bad_sink_formula,
                "test.source.opc_source_overlay",
                case="opc_source_overlay_multi_part_mixed",
                corpus=corpus,
                samples_per_case=2,
                elapsed_samples=[10, 12],
                sample_order=[0, 1],
                source=source,
                sink=sink,
                output_sha256=output_digest,
                allow_pending_hashes=True,
            )

        overflow_overlay = copy.deepcopy(overlay)
        overflow_overlay["sink_max_bytes"] = perf_abba_summary.U64_MAX
        overflow_sink = copy.deepcopy(sink)
        overflow_sink["accepted_bytes"] = perf_abba_summary.U64_MAX
        with self.assertRaisesRegex(
            perf_abba_summary.AbbaSummaryInputError, "formula overflows u64"
        ):
            perf_abba_summary._validate_opc_source_overlay(
                overflow_overlay,
                "test.source.opc_source_overlay",
                case="opc_source_overlay_multi_part_mixed",
                corpus=corpus,
                samples_per_case=2,
                elapsed_samples=[10, 12],
                sample_order=[0, 1],
                source=source,
                sink=overflow_sink,
                output_sha256=output_digest,
                allow_pending_hashes=True,
            )

        with self.assertRaisesRegex(
            perf_abba_summary.AbbaSummaryInputError, "exact sample permutation"
        ):
            perf_abba_summary._validate_opc_source_overlay(
                overlay,
                "test.source.opc_source_overlay",
                case="opc_source_overlay_multi_part_mixed",
                corpus=corpus,
                samples_per_case=2,
                elapsed_samples=[10, 12],
                sample_order=[0, 0],
                source=source,
                sink=sink,
                output_sha256=output_digest,
                allow_pending_hashes=True,
            )

    def test_docx_section_layout_fixed_identity_and_shape_are_exact(self):
        reports = four_legs()
        samples = [20 + index for index in range(15)]
        for leg_index, leg in enumerate(reports):
            leg["configuration"]["cases"] = [perf_abba_summary.DOCX_SECTION_LAYOUT_CASE]
            leg["configuration"]["corpus_shapes"] = ["unlisted-shape"]
            leg["results"] = [docx_section_layout_row(samples, offset=leg_index)]
        summary = perf_abba_summary.summarize_reports(reports)
        self.assertEqual(summary["verification"]["result_count"], 1)

        corpus = docx_section_layout_corpus()
        identity = perf_abba_summary._canonical_json(corpus, "test.corpus")
        indexed = {
            (
                perf_abba_summary.DOCX_SECTION_LAYOUT_CASE,
                identity,
            ): {"elapsed_ns": {"samples": [1] * 15}}
        }
        configuration = {
            "cases": [perf_abba_summary.DOCX_SECTION_LAYOUT_CASE],
            "corpus_shapes": ["unlisted-shape"],
            "samples_per_case": 15,
        }
        perf_abba_summary._validate_configuration_rows(
            configuration, indexed, "test"
        )

        duplicate = copy.deepcopy(corpus)
        duplicate["archive_sha256"] = "c" * 64
        duplicate_identity = perf_abba_summary._canonical_json(duplicate, "test.corpus")
        with self.assertRaisesRegex(
            perf_abba_summary.AbbaSummaryInputError, "exactly one fixed corpus identity"
        ):
            perf_abba_summary._validate_configuration_rows(
                configuration,
                {(
                    perf_abba_summary.DOCX_SECTION_LAYOUT_CASE,
                    identity,
                ): {"elapsed_ns": {"samples": [1] * 15}}, (
                    perf_abba_summary.DOCX_SECTION_LAYOUT_CASE,
                    duplicate_identity,
                ): {"elapsed_ns": {"samples": [1] * 15}}},
                "test",
            )

    def test_docx_section_layout_schema_projection_and_mutations(self):
        samples = [10, 12]
        source = docx_section_layout_source(samples, offset=5)
        corpus = docx_section_layout_corpus()
        overlay = source["docx_section_layout"]
        sink = docx_section_layout_sink()

        def validate(value=overlay, *, source_value=source, sink_value=sink, order=(0, 1)):
            perf_abba_summary._validate_docx_section_layout(
                value,
                "test.source.docx_section_layout",
                corpus=corpus,
                samples_per_case=2,
                elapsed_samples=samples,
                sample_order=list(order),
                source=source_value,
                sink=sink_value,
                output_sha256="5e87e9ca9fd6b9a98933c36d9aee1e848bc97eb7693229f7949514830121c4ec",
                allow_pending_hashes=True,
            )

        validate()
        projected = perf_abba_summary._source_identity_projection(source)
        projected_overlay = projected["docx_section_layout"]
        for field in perf_abba_summary._DOCX_SECTION_LAYOUT_DYNAMIC_FIELDS:
            self.assertNotIn(field, projected_overlay)
        self.assertIn("expected_output_sha256", projected_overlay)
        self.assertIn("output_sha256", projected_overlay)
        for field in perf_abba_summary._DOCX_SECTION_LAYOUT_ROOT_SOURCE_VECTOR_FIELDS:
            self.assertNotIn(field, projected)

        mutations = (
            (lambda value: value.pop("publication_ns"), "schema mismatch"),
            (lambda value: value["publication_ns"].pop(), "exactly 2 samples"),
            (lambda value: value["publication_ns"].__setitem__(1, 7), "bind to sorted elapsed_ns"),
            (lambda value: value.__setitem__("sample_order", [1, 0]), "disagrees with elapsed"),
            (lambda value: value.__setitem__("phase_sum_verified", False), "phase_sum_verified must be true"),
            (lambda value: value.__setitem__("source_sha256", "f" * 64), "source_sha256 disagrees"),
            (lambda value: value.__setitem__("expected_output_sha256", "f" * 64), "expected_output_sha256 disagrees"),
            (lambda value: value.__setitem__("total_main_story_paragraph_count", 256), "total_main_story_paragraph_count"),
        )
        for mutate, message in mutations:
            malformed = copy.deepcopy(overlay)
            mutate(malformed)
            with self.subTest(message=message), self.assertRaisesRegex(
                perf_abba_summary.AbbaSummaryInputError, message
            ):
                validate(malformed)

        bad_source = copy.deepcopy(source)
        bad_source["read_calls"][0] += 1
        with self.assertRaisesRegex(
            perf_abba_summary.AbbaSummaryInputError, "source_read_calls disagrees"
        ):
            validate(source_value=bad_source)

        bad_sink = copy.deepcopy(sink)
        bad_sink["largest_write"] = 65_537
        with self.assertRaisesRegex(
            perf_abba_summary.AbbaSummaryInputError, "largest_write exceeds"
        ):
            validate(sink_value=bad_sink)

        with self.assertRaisesRegex(
            perf_abba_summary.AbbaSummaryInputError, "exact sample permutation"
        ):
            validate(order=(0, 0))

    def test_allocator_binary_identity_is_not_accepted_for_latency_abba(self):
        legs = four_legs()
        for leg in legs:
            leg["tool"]["binary"] = "litchi-perf-baseline-alloc"
        with self.assertRaisesRegex(
            perf_abba_summary.AbbaSummaryInputError,
            "binary.*latency ABBA",
        ):
            perf_abba_summary.summarize_reports(legs)

        legs = four_legs()
        for leg in legs:
            leg["results"][0]["corpus"]["shape"] = "new-shape"
        with self.assertRaisesRegex(
            perf_abba_summary.AbbaSummaryInputError, "cover result shapes"
        ):
            perf_abba_summary.summarize_reports(legs)

        fixed_ods = four_legs()
        for leg in fixed_ods:
            leg["configuration"]["cases"] = [
                "ods_source_backed_one_edit_save",
                "ods_source_backed_one_percent_edit_save",
            ]
            for index, result in enumerate(leg["results"]):
                result["case"] = leg["configuration"]["cases"][index]
                result["corpus"].update(
                    name="ods-media-publication",
                    generator="litchi-ods-media-publication-v1",
                    shape="media-rich",
                )
        perf_abba_summary.summarize_reports(fixed_ods)

        for leg in fixed_ods:
            for result in leg["results"]:
                result["corpus"]["generator"] = "unknown-generator"
        with self.assertRaisesRegex(
            perf_abba_summary.AbbaSummaryInputError, "cover result shapes"
        ):
            perf_abba_summary.summarize_reports(fixed_ods)

        filesystem = four_legs()
        for leg in filesystem:
            leg["configuration"]["cases"] = ["docx_file_source_full_text"]
            leg["configuration"]["filesystem_cache_states"] = ["warm"]
            leg["configuration"]["filesystem_fresh_child_per_sample"] = True
            for result in leg["results"]:
                result["case"] = "docx_file_source_full_text"
                result["corpus"]["shape"] = "media-rich"
            leg["filesystem_evidence"] = [
                {
                    "case": "docx_file_source_full_text",
                    "corpus": copy.deepcopy(leg["results"][0]["corpus"]),
                    "warmup_iterations": 1,
                    "sample_count": 15,
                    "cache_states": ["warm"],
                    "fresh_child_per_sample": True,
                    "samples": [
                        {
                            "sample_index": index,
                            "cache_state": "warm",
                            "elapsed_ns": 100 + index,
                            "parent_wall_ns": 200 + index,
                            "cold_advice": "not_requested",
                            "logical_read_counter_scope": "test_scope",
                            "logical_read_calls": 0,
                            "logical_read_requested_bytes": 0,
                            "logical_read_bytes": 0,
                            "logical_read_largest_requested_bytes": 0,
                            "logical_read_largest_returned_bytes": 0,
                            "max_concurrent_reads": 0,
                            "logical_read_request_sizes": [],
                            "logical_read_request_size_buckets": {
                                "bytes_0": 0,
                                "bytes_1_to_512": 0,
                                "bytes_513_to_4096": 0,
                                "bytes_4097_to_16384": 0,
                                "bytes_16385_to_65536": 0,
                                "bytes_over_65536": 0,
                            },
                            "process_metrics": None,
                            "output_sha256": None,
                            "output_bytes": None,
                            "opc_materialized_parts": None,
                            "cfb_changed_spans": None,
                            "cfb_published_bytes": None,
                        }
                        for index in range(15)
                    ],
                }
            ]
        summary = perf_abba_summary.summarize_reports(filesystem)
        self.assertEqual({result["shape"] for result in summary["results"]}, {"media-rich"})

    def test_source_sink_and_output_identity_statuses_distinguish_absence(self):
        summary = perf_abba_summary.summarize_reports(four_legs())
        self.assertEqual(
            summary["verification"]["source_identity"],
            {"verified_equal": 2, "consistently_absent": 0},
        )
        self.assertEqual(
            summary["verification"]["sink_identity"],
            {"verified_equal": 2, "consistently_absent": 0},
        )
        self.assertFalse(summary["verification"]["output_sha256_identity_verified"])

        absent = four_legs()
        for leg in absent:
            for result in leg["results"]:
                result.pop("source")
                result.pop("sink")
        summary = perf_abba_summary.summarize_reports(absent)
        self.assertEqual(
            summary["verification"]["source_identity"],
            {"verified_equal": 0, "consistently_absent": 2},
        )
        self.assertEqual(
            summary["verification"]["sink_identity"],
            {"verified_equal": 0, "consistently_absent": 2},
        )
        self.assertFalse(summary["verification"]["source_identity_verified"])
        self.assertFalse(summary["verification"]["sink_identity_verified"])

        output_hash = "a" * 64
        with_output = four_legs()
        for leg in with_output:
            for result in leg["results"]:
                result["output_sha256"] = output_hash
        summary = perf_abba_summary.summarize_reports(with_output)
        self.assertEqual(
            summary["verification"]["output_sha256_identity"],
            {"verified_equal": 2, "consistently_absent": 0},
        )
        self.assertTrue(summary["verification"]["output_sha256_identity_verified"])

        malformed = copy.deepcopy(with_output)
        malformed[0]["results"][0]["output_sha256"] = "A" * 64
        with self.assertRaisesRegex(perf_abba_summary.AbbaSummaryInputError, "lowercase"):
            perf_abba_summary.summarize_reports(malformed)
        mismatch = copy.deepcopy(with_output)
        mismatch[2]["results"][0]["output_sha256"] = "b" * 64
        with self.assertRaisesRegex(perf_abba_summary.AbbaSummaryInputError, "output_sha256"):
            perf_abba_summary.summarize_reports(mismatch)
        mixed = copy.deepcopy(with_output)
        mixed[1]["results"][0].pop("output_sha256")
        with self.assertRaisesRegex(perf_abba_summary.AbbaSummaryInputError, "presence"):
            perf_abba_summary.summarize_reports(mixed)

    def test_complete_report_json_is_canonicalized_and_hashed(self):
        legs = four_legs()
        summary = perf_abba_summary.summarize_reports(legs)
        for label, leg in zip(perf_abba_summary.LEG_ORDER, legs):
            canonical = perf_abba_summary._canonical_json(leg, f"{label}.report")
            self.assertEqual(
                summary["report_identity"][label]["canonical_sha256"],
                hashlib.sha256(canonical.encode("utf-8")).hexdigest(),
            )
            self.assertEqual(
                summary["report_identity"][label]["canonical_sha256"],
                perf_abba_summary._canonical_sha256(leg, f"{label}.report"),
            )

        malformed = four_legs()
        malformed[0]["unserializable"] = float("nan")
        with self.assertRaisesRegex(perf_abba_summary.AbbaSummaryInputError, "non-finite"):
            perf_abba_summary.summarize_reports(malformed)
        malformed = four_legs()
        malformed[0][1] = "non-string key"
        with self.assertRaisesRegex(perf_abba_summary.AbbaSummaryInputError, "object key"):
            perf_abba_summary.summarize_reports(malformed)
        with tempfile.TemporaryDirectory() as directory:
            duplicate = Path(directory) / "duplicate.json"
            duplicate.write_text('{"schema_version": 1, "schema_version": 2}', encoding="utf-8")
            with self.assertRaisesRegex(
                perf_abba_summary.AbbaSummaryInputError, "duplicate"
            ):
                perf_abba_summary.load_report(duplicate)

    def test_source_and_sink_identity_mismatches_fail_closed(self):
        for field in ("source", "sink"):
            legs = four_legs()
            legs[2]["results"][0][field] = {"changed": True}
            with self.subTest(field=field), self.assertRaisesRegex(
                perf_abba_summary.AbbaSummaryInputError, f"{field} identity"
            ):
                perf_abba_summary.summarize_reports(legs)

    def test_cfb_phase_timings_are_measurements_not_source_identity(self):
        legs = four_legs()
        for leg_index, leg in enumerate(legs):
            for result in leg["results"]:
                result["source"] = {
                    "cfb_open_stream": {
                        "expected_payload_sha256": "a" * 64,
                        "source_version_check": "stable version fence",
                        "logical_read_calls": [2, 2, 2, 2, 2],
                        "open_ns": [100 + leg_index],
                        "operation_ns": [200 + leg_index],
                        "per_operation_ns": [[200 + leg_index]],
                        "total_ns": [300 + leg_index],
                    }
                }
        summary = perf_abba_summary.summarize_reports(legs)
        source = summary["results"][0]["source"]["cfb_open_stream"]
        self.assertEqual(source["expected_payload_sha256"], "a" * 64)
        self.assertEqual(source["source_version_check"], "stable version fence")
        for field in perf_abba_summary.CFB_OPEN_STREAM_SOURCE_MEASUREMENTS:
            self.assertNotIn(field, source)

        legs[2]["results"][0]["source"]["cfb_open_stream"][
            "source_version_check"
        ] = "changed version fence"
        with self.assertRaisesRegex(
            perf_abba_summary.AbbaSummaryInputError, "source identity"
        ):
            perf_abba_summary.summarize_reports(legs)

    def test_ods_source_cell_timings_and_reads_are_measurements_not_identity(self):
        legs = four_legs()
        for leg_index, leg in enumerate(legs):
            for result in leg["results"]:
                result["source"] = {
                    "read_calls": [10 + leg_index],
                    "read_bytes": [100 + leg_index],
                    "max_in_flight_reads": [],
                    "ods_source_cell": {
                        "source_archive_sha256": "a" * 64,
                        "output_sha256": "b" * 64,
                        "source_hash_verified": True,
                        "lifecycle_ns": [100 + leg_index],
                        "content_source_read_calls": [2 + leg_index],
                        "content_source_read_bytes": [20 + leg_index],
                    },
                }
        summary = perf_abba_summary.summarize_reports(legs)
        source = summary["results"][0]["source"]
        self.assertNotIn("read_calls", source)
        self.assertNotIn("read_bytes", source)
        ods = source["ods_source_cell"]
        self.assertEqual(ods["source_archive_sha256"], "a" * 64)
        self.assertEqual(ods["output_sha256"], "b" * 64)
        for field in perf_abba_summary.ODS_SOURCE_CELL_MEASUREMENTS:
            self.assertNotIn(field, ods)

        legs[2]["results"][0]["source"]["ods_source_cell"][
            "source_archive_sha256"
        ] = "c" * 64
        with self.assertRaisesRegex(
            perf_abba_summary.AbbaSummaryInputError, "source identity"
        ):
            perf_abba_summary.summarize_reports(legs)

    def test_xlsx_cell_value_timings_are_measurements_not_source_identity(self):
        legs = four_legs()
        for leg_index, leg in enumerate(legs):
            for result in leg["results"]:
                result["source"] = {
                    "read_calls": [10],
                    "read_bytes": [100],
                    "xlsx_cell_values": {
                        "source_archive_sha256": "a" * 64,
                        "output_sha256": ["b" * 64],
                        "semantic_sha256": ["c" * 64],
                        "open_ns": [100 + leg_index],
                        "plan_ns": [200 + leg_index],
                        "commit_ns": [300 + leg_index],
                        "publication_ns": [400 + leg_index],
                        "reopen_ns": [500 + leg_index],
                    },
                }
        summary = perf_abba_summary.summarize_reports(legs)
        source = summary["results"][0]["source"]
        self.assertEqual(source["read_calls"], [10])
        xlsx = source["xlsx_cell_values"]
        self.assertEqual(xlsx["source_archive_sha256"], "a" * 64)
        self.assertEqual(xlsx["output_sha256"], ["b" * 64])
        for field in perf_abba_summary.XLSX_CELL_VALUES_SOURCE_MEASUREMENTS:
            self.assertNotIn(field, xlsx)

        legs[2]["results"][0]["source"]["read_calls"] = [11]
        with self.assertRaisesRegex(
            perf_abba_summary.AbbaSummaryInputError, "source identity"
        ):
            perf_abba_summary.summarize_reports(legs)

    def test_xls_numeric_operation_shape_projection_is_narrow_and_exact(self):
        legs = xls_numeric_legs()
        summary = perf_abba_summary.summarize_reports(legs)
        numeric_result = next(
            result for result in summary["results"] if "xls_numeric" in result["source"]
        )
        projected = numeric_result["source"]["xls_numeric"][
            "operation_evidence"
        ][0]
        self.assertNotIn("source_mode", projected)
        for field in perf_abba_summary.XLS_NUMERIC_TIMING_VECTOR_FIELDS:
            self.assertNotIn(field, numeric_result["source"]["xls_numeric"])
        for field in perf_abba_summary.XLS_NUMERIC_OPERATION_MEASUREMENT_FIELDS:
            self.assertNotIn(field, projected)
        for field in (
            "counter_scope",
            "source_bytes",
            "fingerprint_chunk_bytes",
            "publication_chunk_bytes",
            "fingerprint_buffer_bytes",
            "publication_buffer_bytes",
            "candidate_reopen_logical_artifact_bytes",
            "selected_stream_logical_bytes",
            "splice_count",
            "changed_span_count",
            "replacement_bytes",
            "target_materialization_vec_bytes",
            "target_materialization_clone_bytes",
            "publication_write_calls",
            "atomic_save_event_scope",
        ):
            self.assertIn(field, projected)
        raw_hashes = {
            perf_abba_summary._canonical_sha256(leg, "xls") for leg in legs
        }
        self.assertGreater(len(raw_hashes), 1)

        malformed = copy.deepcopy(legs)
        malformed[0]["results"][0]["source"]["xls_numeric"][
            "operation_evidence"
        ][0]["planning_fingerprint_scans"] = 1
        with self.assertRaisesRegex(
            perf_abba_summary.AbbaSummaryInputError, "source-mode policy"
        ):
            perf_abba_summary.summarize_reports(malformed)

        malformed = copy.deepcopy(legs)
        malformed[0]["results"][0]["source"]["xls_numeric"][
            "operation_evidence"
        ][0].pop("splice_count")
        with self.assertRaisesRegex(
            perf_abba_summary.AbbaSummaryInputError, "keys mismatch"
        ):
            perf_abba_summary.summarize_reports(malformed)

        malformed = copy.deepcopy(legs)
        malformed[2]["results"][0]["source"]["xls_numeric"][
            "operation_evidence"
        ][0]["counter_scope"] = "forged scope"
        with self.assertRaisesRegex(
            perf_abba_summary.AbbaSummaryInputError, "validated CFB scope"
        ):
            perf_abba_summary.summarize_reports(malformed)

        malformed = copy.deepcopy(legs)
        malformed[1]["results"][0]["source"]["xls_numeric"][
            "operation_evidence"
        ].pop()
        with self.assertRaisesRegex(
            perf_abba_summary.AbbaSummaryInputError, "cardinality"
        ):
            perf_abba_summary.summarize_reports(malformed)

    def test_xls_numeric_current_evidence_discriminator_is_fail_closed(self):
        missing_current = xls_numeric_legs()
        numeric = missing_current[0]["results"][0]["source"]["xls_numeric"]
        numeric.pop("operation_evidence")
        with self.assertRaisesRegex(
            perf_abba_summary.AbbaSummaryInputError, "requires operation_evidence"
        ):
            perf_abba_summary.summarize_reports(missing_current)

        scope_downgrade = xls_numeric_legs()
        for leg in scope_downgrade:
            numeric = leg["results"][0]["source"]["xls_numeric"]
            numeric.pop("operation_evidence")
            numeric.pop("operation_evidence_schema")
            numeric[
                "source_counter_scope"
            ] = perf_abba_summary.XLS_NUMERIC_LEGACY_SOURCE_COUNTER_SCOPE
        with self.assertRaisesRegex(
            perf_abba_summary.AbbaSummaryInputError, "current operation-evidence discriminator"
        ):
            perf_abba_summary.summarize_reports(scope_downgrade)

        missing_schema = xls_numeric_legs()
        for leg in missing_schema:
            numeric = leg["results"][0]["source"]["xls_numeric"]
            numeric.pop("operation_evidence_schema")
        with self.assertRaisesRegex(
            perf_abba_summary.AbbaSummaryInputError, "requires operation_evidence_schema"
        ):
            perf_abba_summary.summarize_reports(missing_schema)

        eager = xls_numeric_legs(
            case="xls_numeric_eager_number_edit_save", implementation="eager"
        )
        numeric = eager[0]["results"][0]["source"]["xls_numeric"]
        numeric["operation_evidence_schema"] = (
            perf_abba_summary.XLS_NUMERIC_OPERATION_EVIDENCE_SCHEMA
        )
        numeric["operation_evidence"] = []
        with self.assertRaisesRegex(
            perf_abba_summary.AbbaSummaryInputError, "eager implementation"
        ):
            perf_abba_summary.summarize_reports(eager)

        plan = xls_numeric_legs(
            case="xls_numeric_plan_only_number_edit_save", implementation="plan_only"
        )
        plan_numeric = plan[0]["results"][0]["source"]["xls_numeric"]
        plan_numeric["target_artifact_materialized_at_commit"] = True
        with self.assertRaisesRegex(
            perf_abba_summary.AbbaSummaryInputError, "implementation"
        ):
            perf_abba_summary.summarize_reports(plan)

    def test_xls_numeric_timing_vectors_bind_to_phases_and_elapsed_order(self):
        malformed = xls_numeric_legs()
        malformed[0]["results"][0]["source"]["xls_numeric"]["total_ns"][0] += 1
        with self.assertRaisesRegex(
            perf_abba_summary.AbbaSummaryInputError, "per-sample phase sum"
        ):
            perf_abba_summary.summarize_reports(malformed)

        malformed = xls_numeric_legs()
        malformed[0]["results"][0]["elapsed_ns"]["sample_order"] = [0] * 15
        with self.assertRaisesRegex(
            perf_abba_summary.AbbaSummaryInputError, "complete permutation"
        ):
            perf_abba_summary.summarize_reports(malformed)

        malformed = xls_numeric_legs()
        malformed[0]["results"][0]["elapsed_ns"]["sample_order"][:2] = [1, 0]
        with self.assertRaisesRegex(
            perf_abba_summary.AbbaSummaryInputError, "bind to elapsed_ns.samples"
        ):
            perf_abba_summary.summarize_reports(malformed)

    def test_xls_numeric_source_mode_roles_are_bound_to_abba_leg(self):
        malformed = xls_numeric_legs()
        malformed[0]["results"][0]["source"]["xls_numeric"][
            "operation_evidence"
        ][0]["source_mode"] = "owned_immutable_arc"
        with self.assertRaisesRegex(
            perf_abba_summary.AbbaSummaryInputError, "ABBA implementation-role"
        ):
            perf_abba_summary.summarize_reports(malformed)

        malformed = xls_numeric_legs()
        malformed[2]["results"][0]["source"]["xls_numeric"][
            "operation_evidence"
        ][0]["source_mode"] = "generic_read_at"
        with self.assertRaisesRegex(
            perf_abba_summary.AbbaSummaryInputError, "ABBA implementation-role"
        ):
            perf_abba_summary.summarize_reports(malformed)

    def test_xls_projection_separates_abba_role_from_diagnostic_location(self):
        reports = xls_numeric_legs()
        projections = {
            role: perf_abba_summary._project_report(
                report,
                f"package/abba-0268/{role}",
                profile=perf_abba_summary.detect_report_profile(report, role),
                report_role=role,
            )
            for role, report in zip(perf_abba_summary.LEG_ORDER, reports)
        }
        self.assertEqual(
            perf_abba_summary._summarize_projected_reports(projections),
            perf_abba_summary.summarize_reports(reports),
        )

        malformed = copy.deepcopy(reports[0])
        malformed["results"][0]["source"]["xls_numeric"]["operation_evidence"][0][
            "source_mode"
        ] = "owned_immutable_arc"
        with self.assertRaisesRegex(
            perf_abba_summary.AbbaSummaryInputError,
            r"package/abba-0268/a1.*ABBA implementation-role",
        ):
            perf_abba_summary._project_report(
                malformed,
                "package/abba-0268/a1",
                profile=perf_abba_summary.detect_report_profile(malformed, "a1"),
                report_role="a1",
            )

    def test_xls_numeric_top_level_contract_tampering_is_rejected(self):
        for field, value, message in (
            ("family", "other", "family"),
            ("update_count", 0, "update_count"),
            ("sink_capacity_bytes", 1, "sink_capacity_bytes"),
            ("expected_output_sha256", "bad", "lowercase"),
            ("sample_count", 14, "sample_count|elapsed_ns.samples"),
        ):
            malformed = xls_numeric_legs()
            malformed[0]["results"][0]["source"]["xls_numeric"][field] = value
            with self.subTest(field=field), self.assertRaisesRegex(
                perf_abba_summary.AbbaSummaryInputError, message
            ):
                perf_abba_summary.summarize_reports(malformed)

        malformed = xls_numeric_legs()
        malformed[0]["results"][0]["source"]["xls_numeric"]["sink_digests"][0] = "e" * 64
        with self.assertRaisesRegex(
            perf_abba_summary.AbbaSummaryInputError, "sink_digests"
        ):
            perf_abba_summary.summarize_reports(malformed)

        malformed = xls_numeric_legs()
        malformed[0]["results"][0]["source"]["xls_numeric"]["unexpected"] = True
        with self.assertRaisesRegex(
            perf_abba_summary.AbbaSummaryInputError, "unknown keys"
        ):
            perf_abba_summary.summarize_reports(malformed)

        for field, value, message in (
            ("archive_bytes", 131073, "corpus.archive_bytes"),
            ("archive_bytes", perf_abba_summary.XLS_NUMERIC_CFB_MAX_BYTES + 1, "2 GiB"),
        ):
            malformed = xls_numeric_legs()
            malformed[0]["results"][0]["corpus"][field] = value
            with self.subTest(field=field, value=value), self.assertRaisesRegex(
                perf_abba_summary.AbbaSummaryInputError,
                f"{message}|native corpus manifest",
            ):
                perf_abba_summary.summarize_reports(malformed)

        malformed = xls_numeric_legs()
        malformed[0]["results"][0]["source"]["xls_numeric"]["output_cfb_bytes"] += 1
        with self.assertRaisesRegex(
            perf_abba_summary.AbbaSummaryInputError, "output_cfb_bytes"
        ):
            perf_abba_summary.summarize_reports(malformed)

        malformed = xls_numeric_legs()
        malformed[0]["results"][0]["output_sha256"] = "e" * 64
        with self.assertRaisesRegex(
            perf_abba_summary.AbbaSummaryInputError, "expected_output_sha256|output_sha256"
        ):
            perf_abba_summary.summarize_reports(malformed)

        malformed = xls_numeric_legs()
        malformed[0]["results"][0]["sink"]["write_calls"] = 3
        with self.assertRaisesRegex(
            perf_abba_summary.AbbaSummaryInputError, "sink.write_calls"
        ):
            perf_abba_summary.summarize_reports(malformed)

    def test_xls_numeric_native_source_presence_and_identity_are_required(self):
        for mutation in ("null", "omitted", "numeric_omitted"):
            malformed = xls_numeric_legs()
            source = malformed[0]["results"][0]["source"]
            if mutation == "null":
                malformed[0]["results"][0]["source"] = None
            elif mutation == "omitted":
                malformed[0]["results"][0].pop("source")
            else:
                source.pop("xls_numeric")
            with self.subTest(mutation=mutation), self.assertRaisesRegex(
                perf_abba_summary.AbbaSummaryInputError,
                "source(?:\\.xls_numeric)? is required",
            ):
                perf_abba_summary.summarize_reports(malformed)

        malformed = xls_numeric_legs()
        malformed[1]["results"][0]["source"]["read_calls"][0] += 1
        with self.assertRaisesRegex(
            perf_abba_summary.AbbaSummaryInputError, "source identity"
        ):
            perf_abba_summary.summarize_reports(malformed)

    def test_xls_numeric_native_source_ingress_vectors_are_exact_raw_and_projected(self):
        mutations = (
            ("calls_omitted", lambda source, _contract: source.pop("read_calls"), "source.read_calls"),
            ("bytes_omitted", lambda source, _contract: source.pop("read_bytes"), "source.read_bytes"),
            ("calls_type", lambda source, _contract: source.update(read_calls={}), "source.read_calls"),
            ("bytes_type", lambda source, _contract: source.update(read_bytes="not-a-vector"), "source.read_bytes"),
            ("calls_cardinality", lambda source, contract: source.update(read_calls=[contract["source_read_calls"]] * 14), "source.read_calls"),
            ("bytes_cardinality", lambda source, contract: source.update(read_bytes=[contract["source_read_bytes"]] * 14), "source.read_bytes"),
            ("calls_zero", lambda source, contract: source.update(read_calls=[0] * 15), "source.read_calls must be positive"),
            ("calls_wrong", lambda source, contract: source.update(read_calls=[contract["source_read_calls"] + 1] * 15), "source.read_calls differs"),
            ("bytes_wrong", lambda source, contract: source.update(read_bytes=[contract["source_read_bytes"] + 1] * 15), "source.read_bytes"),
        )
        for case, (implementation, family, _updates) in perf_abba_summary.XLS_NUMERIC_CASE_CONTRACTS.items():
            contract = perf_abba_summary.XLS_NUMERIC_OUTPUT_CONTRACTS[family]
            valid = xls_numeric_legs(case, implementation, family)
            projections = {
                label: perf_abba_summary._project_report(
                    leg,
                    label,
                    profile=perf_abba_summary.detect_report_profile(leg, label),
                )
                for label, leg in zip(perf_abba_summary.LEG_ORDER, valid)
            }
            projected_summary = perf_abba_summary._summarize_projected_reports(projections)
            self.assertTrue(projected_summary["verification"]["source_identity_verified"])
            for mutation, mutate, message in mutations:
                malformed = copy.deepcopy(valid)
                for leg in malformed:
                    mutate(leg["results"][0]["source"], contract)
                with self.subTest(case=case, mutation=mutation, path="raw"), self.assertRaisesRegex(
                    perf_abba_summary.AbbaSummaryInputError, message
                ):
                    perf_abba_summary.summarize_reports(malformed)
                with self.subTest(case=case, mutation=mutation, path="projected"), self.assertRaisesRegex(
                    perf_abba_summary.AbbaSummaryInputError, message
                ):
                    malformed_projections = {
                        label: perf_abba_summary._project_report(
                            leg,
                            label,
                            profile=perf_abba_summary.detect_report_profile(leg, label),
                        )
                        for label, leg in zip(perf_abba_summary.LEG_ORDER, malformed)
                    }
                    perf_abba_summary._summarize_projected_reports(malformed_projections)

    def test_xls_numeric_owned_input_scope_and_corpus_manifest_are_exact(self):
        malformed = xls_numeric_legs()
        malformed[0]["results"][0]["source"]["xls_numeric"][
            "owned_input_scope"
        ] = "forged complete-bytes claim"
        with self.assertRaisesRegex(
            perf_abba_summary.AbbaSummaryInputError, "owned_input_scope"
        ):
            perf_abba_summary.summarize_reports(malformed)

        for field, expected in perf_abba_summary.XLS_NUMERIC_CORPUS_CONTRACTS[
            "Number"
        ].items():
            malformed = xls_numeric_legs()
            value = malformed[0]["results"][0]["corpus"][field]
            if isinstance(value, bool):
                replacement = not value
            elif isinstance(value, int):
                replacement = value + 1
            elif value is None:
                replacement = {}
            else:
                replacement = f"forged-{value}"
            malformed[0]["results"][0]["corpus"][field] = replacement
            with self.subTest(field=field), self.assertRaisesRegex(
                perf_abba_summary.AbbaSummaryInputError, "native corpus manifest"
            ):
                perf_abba_summary.summarize_reports(malformed)

        renamed = xls_numeric_legs()
        for leg in renamed:
            leg["configuration"]["cases"] = [
                "xls_numeric_source_backed_rk_mulrk_edit_save"
            ]
            leg["results"][0]["case"] = (
                "xls_numeric_source_backed_rk_mulrk_edit_save"
            )
        with self.assertRaisesRegex(
            perf_abba_summary.AbbaSummaryInputError,
            "family|selector|native corpus manifest",
        ):
            perf_abba_summary.summarize_reports(renamed)

        downgraded = xls_numeric_legs()
        for leg in downgraded:
            leg["configuration"]["cases"] = ["synthetic_case"]
            leg["results"][0]["case"] = "synthetic_case"
            leg["results"][0].pop("source")
        with self.assertRaisesRegex(
            perf_abba_summary.AbbaSummaryInputError, "native XLS numeric corpus"
        ):
            perf_abba_summary.summarize_reports(downgraded)

    def test_xls_numeric_case_contracts_cover_all_six_selectors(self):
        for case, (implementation, family, _updates) in (
            perf_abba_summary.XLS_NUMERIC_CASE_CONTRACTS.items()
        ):
            with self.subTest(case=case):
                summary = perf_abba_summary.summarize_reports(
                    xls_numeric_legs(case, implementation, family)
                )
                self.assertIs(summary["verification"]["source_identity_verified"], True)

    def test_xls_numeric_emitted_six_selector_reports_summarize_directly(self):
        reports = xls_numeric_all_selector_legs()
        summary = perf_abba_summary.summarize_reports(reports)
        self.assertEqual(
            {result["case"] for result in summary["results"]},
            set(perf_abba_summary.XLS_NUMERIC_CASE_CONTRACTS),
        )
        self.assertTrue(summary["verification"]["source_identity_verified"])
        self.assertEqual(
            summary["verification"]["result_count"],
            len(perf_abba_summary.XLS_NUMERIC_CASE_CONTRACTS),
        )
        report_hashes = {
            item["canonical_sha256"]
            for item in summary["report_identity"].values()
        }
        self.assertEqual(len(report_hashes), 4)
        self.assertEqual(
            set(summary["report_identity"]),
            set(perf_abba_summary.LEG_ORDER),
        )
        self.assertTrue(summary["verification"]["case_corpus_identity_verified"])
        self.assertTrue(summary["verification"]["output_sha256_identity_verified"])
        self.assertNotEqual(
            summary["implementation_identity"]["control"]["binary_sha256"],
            summary["implementation_identity"]["candidate"]["binary_sha256"],
        )
        for implementation in ("control", "candidate"):
            self.assertIn("binary_identity", summary["implementation_identity"][implementation])

    def test_xls_numeric_family_contracts_reject_synchronized_forgery(self):
        for family, case in (
            ("Number", "xls_numeric_source_backed_number_edit_save"),
            ("RK+MulRK", "xls_numeric_source_backed_rk_mulrk_edit_save"),
        ):
            with self.subTest(family=family):
                contract = perf_abba_summary.XLS_NUMERIC_OUTPUT_CONTRACTS[family]
                malformed = xls_numeric_legs(case, "source_backed", family)
                forged_output = "0" * 64
                for leg in malformed:
                    result = leg["results"][0]
                    numeric = result["source"]["xls_numeric"]
                    numeric["expected_output_sha256"] = forged_output
                    numeric["replacement_bytes"] = [
                        contract["replacement_bytes"] + 1
                    ] * 15
                    numeric["changed_spans"] = [contract["changed_spans"] + 1] * 15
                    numeric["source_fingerprints"] = [forged_output] * 15
                    numeric["target_fingerprints"] = [forged_output] * 15
                    numeric["sink_write_calls"] = [
                        contract["sink_write_calls"] + 1
                    ] * 15
                    numeric["sink_digests"] = [forged_output] * 15
                    result["output_sha256"] = forged_output
                    result["sink"]["write_calls"] = contract["sink_write_calls"] + 1
                    result["source"]["read_calls"] = [
                        contract["sink_write_calls"] + 1
                    ] * 15
                    for operation in numeric["operation_evidence"]:
                        operation["replacement_bytes"] = (
                            contract["replacement_bytes"] + 1
                        )
                        operation["changed_span_count"] = contract["changed_spans"] + 1
                        operation["publication_write_calls"] = (
                            contract["sink_write_calls"] + 1
                        )
                with self.assertRaisesRegex(
                    perf_abba_summary.AbbaSummaryInputError,
                    "exact|native",
                ):
                    perf_abba_summary.summarize_reports(malformed)

    def test_xls_numeric_selector_rename_changes_raw_identity_and_is_rejected(self):
        reports = xls_numeric_all_selector_legs()
        original_hashes = [
            perf_abba_summary._canonical_sha256(report, "native")
            for report in reports
        ]
        renamed = copy.deepcopy(reports)
        renamed_case = "xls_numeric_source_backed_number_edit_save_renamed"
        for report in renamed:
            report["configuration"]["cases"] = [renamed_case] + [
                case
                for case in report["configuration"]["cases"][1:]
            ]
            report["results"][0]["case"] = renamed_case
        renamed_hashes = [
            perf_abba_summary._canonical_sha256(report, "renamed")
            for report in renamed
        ]
        self.assertNotEqual(original_hashes, renamed_hashes)
        with self.assertRaisesRegex(
            perf_abba_summary.AbbaSummaryInputError,
            "native XLS numeric corpus|selector",
        ):
            perf_abba_summary.summarize_reports(renamed)

    def test_tool_configuration_and_result_identity_mismatches_fail_closed(self):
        for mutation, message in (
            (lambda legs: legs[1]["tool"].update(version="other"), "tool identity"),
            (lambda legs: legs[1]["configuration"].update(samples_per_case=6), "configuration"),
            (
                lambda legs: legs[1]["results"].__setitem__(
                    0,
                    {
                        **row("tiny", [1, 1, 1, 1, 1]),
                        "corpus": {
                            **row("tiny", [1, 1, 1, 1, 1])["corpus"],
                            "archive_sha256": "changed" * 10 + "changed"[:4],
                        },
                    },
                ),
                "case/corpus",
            ),
        ):
            legs = four_legs()
            mutation(legs)
            with self.subTest(message=message), self.assertRaisesRegex(
                perf_abba_summary.AbbaSummaryInputError, message
            ):
                perf_abba_summary.summarize_reports(legs)

    def test_reported_statistics_and_sample_cardinality_are_verified(self):
        legs = four_legs()
        legs[0]["results"][0]["elapsed_ns"]["p95"] += 1
        with self.assertRaisesRegex(perf_abba_summary.AbbaSummaryInputError, "disagrees"):
            perf_abba_summary.summarize_reports(legs)

        legs = four_legs()
        legs[3]["results"][0]["elapsed_ns"] = elapsed([1, 2, 3, 4])
        with self.assertRaisesRegex(
            perf_abba_summary.AbbaSummaryInputError, "samples_per_case|sample counts"
        ):
            perf_abba_summary.summarize_reports(legs)

    def test_schema_mapping_and_ceiling_validation_fail_closed(self):
        legs = four_legs()
        legs[2]["schema_version"] = 2
        with self.assertRaisesRegex(perf_abba_summary.AbbaSummaryInputError, "schema_version"):
            perf_abba_summary.summarize_reports(legs)

        with self.assertRaisesRegex(perf_abba_summary.AbbaSummaryInputError, "non-negative"):
            perf_abba_summary.summarize_reports(
                {
                    label.upper(): value
                    for label, value in zip(("a1", "b1", "b2", "a2"), four_legs())
                },
                drift_ceilings={"p50": -1, "mean": 5, "p95": 10, "p99": 15},
            )

    def test_selectors_filter_after_full_identity_validation(self):
        summary = perf_abba_summary.summarize_reports(four_legs(), shapes=["tiny"])
        self.assertEqual(len(summary["results"]), 1)
        self.assertEqual(summary["results"][0]["shape"], "tiny")
        with self.assertRaisesRegex(perf_abba_summary.AbbaSummaryInputError, "did not match"):
            perf_abba_summary.summarize_reports(four_legs(), cases=["missing"])

    def test_output_is_deterministic_and_cli_accepts_positional_reports(self):
        legs = four_legs()
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            paths = []
            for index, leg in enumerate(legs):
                path = root / f"leg-{index}.json"
                path.write_text(json.dumps(leg, sort_keys=False), encoding="utf-8")
                paths.append(path)
            output = root / "summary.json"
            with contextlib.redirect_stdout(io.StringIO()):
                self.assertEqual(
                    perf_abba_summary.main(
                        [*(str(path) for path in paths), "--json-out", str(output)]
                    ),
                    0,
                )
            from_file = json.loads(output.read_text(encoding="utf-8"))
        direct = perf_abba_summary.summarize_reports(legs)
        self.assertEqual(
            json.dumps(from_file, sort_keys=True, separators=(",", ":")),
            json.dumps(direct, sort_keys=True, separators=(",", ":")),
        )


class ProjectedReportSummaryTests(unittest.TestCase):
    def legacy_legs(self):
        legs = four_legs()
        for leg in legs:
            leg["tool"].pop("binary")
            leg["tool"].pop("instrumentation")
            leg.pop("binary_identity")
        return legs

    def projected(self, legs):
        return {
            label: perf_abba_summary._project_report(
                leg,
                label,
                profile=perf_abba_summary.detect_report_profile(leg, label),
            )
            for label, leg in zip(perf_abba_summary.LEG_ORDER, legs)
        }

    def test_legacy_profile_projection_matches_direct_summary(self):
        legs = self.legacy_legs()
        self.assertEqual(
            perf_abba_summary.detect_reports_profile(legs),
            perf_abba_summary.REPORT_PROFILE_LEGACY,
        )
        direct = perf_abba_summary.summarize_reports(legs)
        projected = perf_abba_summary._summarize_projected_reports(self.projected(legs))
        self.assertEqual(projected, direct)
        self.assertNotIn("binary_identity_verified", direct["verification"])

    def test_current_profile_projection_matches_direct_summary(self):
        legs = four_legs()
        self.assertEqual(
            perf_abba_summary.detect_reports_profile(legs),
            perf_abba_summary.REPORT_PROFILE_CURRENT,
        )
        direct = perf_abba_summary.summarize_reports(legs)
        projected = perf_abba_summary._summarize_projected_reports(self.projected(legs))
        self.assertEqual(projected, direct)
        self.assertTrue(projected["verification"]["binary_identity_verified"])

    def test_xls_numeric_projected_summary_matches_direct_summary(self):
        legs = xls_numeric_legs()
        direct = perf_abba_summary.summarize_reports(legs)
        projected = perf_abba_summary._summarize_projected_reports(self.projected(legs))
        self.assertEqual(projected, direct)

    def test_mixed_profile_is_rejected_from_all_raw_legs(self):
        legs = self.legacy_legs()
        legs[0]["tool"].update(binary=perf_abba_summary.HARNESS_TOOL_NAME, instrumentation="none")
        legs[0]["binary_identity"] = copy.deepcopy(four_legs()[0]["binary_identity"])
        with self.assertRaisesRegex(
            perf_abba_summary.AbbaSummaryInputError, "mixed"
        ):
            perf_abba_summary.detect_reports_profile(legs)
        with self.assertRaisesRegex(
            perf_abba_summary.AbbaSummaryInputError, "mixed"
        ):
            perf_abba_summary.summarize_reports(legs)

    def test_projection_rejects_raw_identity_mismatch(self):
        legs = four_legs()
        legs[1]["results"][0]["sink"]["accepted_bytes"] += 1
        with self.assertRaisesRegex(
            perf_abba_summary.AbbaSummaryInputError, "sink identity"
        ):
            perf_abba_summary._summarize_projected_reports(self.projected(legs))

    def test_raw_summary_never_trusts_projection_only_markers(self):
        legs = four_legs()
        row = legs[0]["results"][0]
        row["elapsed_ns"]["samples"][-1] += 1
        # These keys are intentionally accepted only on the internal
        # validated projection path.  A raw caller cannot use them to bypass
        # sample/statistic validation.
        row["_elapsed_statistics"] = {"sample_count": 5, "p50": 999}
        row["_operation_metrics_identity"] = "forged"
        with self.assertRaises(perf_abba_summary.AbbaSummaryInputError):
            perf_abba_summary.summarize_reports(legs)

    def test_projection_helpers_are_private_and_mutation_is_rejected(self):
        self.assertFalse(hasattr(perf_abba_summary, "project_report"))
        self.assertFalse(hasattr(perf_abba_summary, "summarize_projected_reports"))
        projections = self.projected(four_legs())
        for projection in projections.values():
            row = next(iter(projection["results"].values()))
            row["_elapsed_statistics"]["p50"] += 1
        with self.assertRaisesRegex(
            perf_abba_summary.AbbaSummaryInputError, "mutated"
        ):
            perf_abba_summary._summarize_projected_reports(projections)

        forged = {
            label: dict(projection)
            for label, projection in self.projected(four_legs()).items()
        }
        with self.assertRaisesRegex(
            perf_abba_summary.AbbaSummaryInputError, "private validation provenance"
        ):
            perf_abba_summary._summarize_projected_reports(forged)
        with self.assertRaises(TypeError):
            perf_abba_summary.summarize_reports(_validated=("current-v1", {}))


if __name__ == "__main__":
    unittest.main()
