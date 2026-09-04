import copy
import hashlib
import json
import math
import subprocess
import tempfile
import unittest
from pathlib import Path

from tools import perf_abba_package
from tools import perf_compare


TOOL = {
    "name": "litchi-perf-baseline",
    "version": "0.1.0",
    "binary": "litchi-perf-baseline",
    "profile": "release",
    "target_os": "linux",
    "target_arch": "x86_64",
    "instrumentation": "none",
}


def policy():
    corpus = report()["results"][0]["corpus"]
    corpus_identity = json.dumps(
        corpus, sort_keys=True, separators=(",", ":"), allow_nan=False
    )
    return {
        "schema_version": 2,
        "policy_id": "test-policy-v2",
        "minimum_samples": 5,
        "expected_result_count": 1,
        "expected_result_keys_sha256": perf_compare.result_key_manifest_sha256(
            [("opc_open", corpus_identity)]
        ),
        "required_cases": ["opc_open"],
        "require_clean_worktree": True,
        "require_distinct_revisions": True,
        "require_binary_identity": True,
        "tool_identity": copy.deepcopy(TOOL),
        "build_identity_fields": [
            "rustc_version",
            "allocator",
            "rustflags",
            "cargo_build_target",
            "logical_cpus_available",
        ],
        "nullable_build_identity_fields": [],
        "expected_configuration": {"samples_per_case": 5},
        "latency_thresholds_percent": {"p50": 5.0, "p95": 8.0, "p99": 12.0},
        "metric_classes": [
            {
                "name": "allocation",
                "max_regression_percent": 3.0,
                "path_globs": ["**/allocation_calls", "**/allocation_calls/values"],
                "presence": "optional",
            },
            {
                "name": "rss",
                "max_regression_percent": 4.0,
                "path_globs": ["**/*rss_bytes"],
                "presence": "optional",
            },
            {
                "name": "work",
                "max_regression_percent": 5.0,
                "path_globs": ["source/read_calls", "**/instructions"],
                "presence": "optional",
            },
        ],
    }


def report(value=100, revision="baseline", corpus_sha="abc"):
    samples = [value] * 5
    candidate = revision != "baseline"
    return {
        "schema_version": 1,
        "tool": copy.deepcopy(TOOL),
        "binary_identity": {
            "path": "/tmp/litchi-perf-candidate" if candidate else "/tmp/litchi-perf-baseline",
            "binary_sha256": "b" * 64 if candidate else "a" * 64,
            "binary_bytes": 200 if candidate else 100,
            "mode_bits": 0o755,
            "executable": True,
            "profile": "release",
        },
        "environment": {
            "rustc_version": "rustc 1.95.0",
            "git_revision": revision,
            "git_worktree_dirty": False,
            "logical_cpus_available": 2,
            "allocator": "system",
            "rustflags": "-C target-cpu=x86-64-v3",
            "cargo_build_target": "x86_64-unknown-linux-gnu",
            "os": "linux",
            "kernel": "Linux 6.12",
            "cpu_model": "Test CPU",
            "total_memory_bytes": 16_000_000_000,
            "page_size_bytes": 4096,
        },
        "configuration": {
            "samples_per_case": 5,
            "cases": ["opc_open"],
            "execution_workers": [1, 2],
        },
        "results": [
            {
                "case": "opc_open",
                "corpus": {
                    "name": "opc-fixed",
                    "generator": "opc-v1",
                    "shape": "tiny",
                    "payload_kind": "compressible",
                    "archive_sha256": corpus_sha,
                },
                "elapsed_ns": {
                    "unit": "ns",
                    "samples": samples,
                "p50": value,
                "p95": value,
                "p99": value,
            },
                "metrics": {
                    "allocation_calls": [1000] * 5,
                    "peak_rss_bytes": [10_000] * 5,
                    "instructions": [50_000] * 5,
                },
                "source": {"read_calls": [10] * 5},
            }
        ],
    }


def managed_xlsx_reports():
    case = "xlsx_source_backed_managed_cell_values_one_edit_save"
    payload_limit = 100
    headroom = 64 * 1024
    baseline = report()
    current = report(revision="current")
    for item in (baseline, current):
        item["configuration"]["cases"] = [case]
        item["configuration"][
            "xlsx_cell_values_managed_planning_memory_headroom"
        ] = headroom
        item["results"][0]["case"] = case
        item["results"][0]["source"]["xlsx_cell_values"] = {
            "cache_budget_managed": True,
            "payload_memory_limit": payload_limit,
            "publication_planning_memory_headroom": headroom,
            "cache_budget_memory_limit": payload_limit + headroom,
        }
    comparison_policy = policy()
    comparison_policy["required_cases"] = [case]
    corpus_identity = json.dumps(
        baseline["results"][0]["corpus"],
        sort_keys=True,
        separators=(",", ":"),
        allow_nan=False,
    )
    comparison_policy["expected_result_keys_sha256"] = (
        perf_compare.result_key_manifest_sha256([(case, corpus_identity)])
    )
    return baseline, current, comparison_policy


def legacy_managed_xlsx_reports():
    baseline, current, comparison_policy = managed_xlsx_reports()
    for item in (baseline, current):
        item["configuration"].pop(
            "xlsx_cell_values_managed_planning_memory_headroom"
        )
        evidence = item["results"][0]["source"]["xlsx_cell_values"]
        evidence.pop("payload_memory_limit")
        evidence.pop("publication_planning_memory_headroom")
        evidence["cache_budget_memory_limit"] = 100
    return baseline, current, comparison_policy


def metric_vector(values, *, status="measured", scope="test_scope"):
    wrapper = {"status": status, "scope": scope}
    if values is not None:
        wrapper["values"] = values
    return wrapper


def pattern_vector(values, *, status="measured", scope="pattern_scope"):
    return metric_vector(values, status=status, scope=scope)


def operation_metrics_report_fields():
    def not_applicable(scope):
        return metric_vector(None, status="not_applicable", scope=scope)
    process_scope = "procfs_operation_delta"
    proc_io_scope = "child_process_interval_delta_including_procfs_probe_overhead"
    source_scope = "operation_logical_read_at"
    pattern_scope = "operation_logical_read_at_range_order_not_physical_io"
    cfb_source_scope = "timed_cfb_phase_logical_read_at"
    return {
        "sample_count": 5,
        "sample_indices": list(range(5)),
        "alignment": perf_compare.OPERATION_ALIGNMENT,
        "latency_claim": perf_compare.EVIDENCE_ONLY_LATENCY_CLAIM,
        "source": {
            "status": "measured",
            "counter_scope": "timed_read_at",
            "logical_read_calls": metric_vector([10] * 5, scope=source_scope),
            "logical_read_requested_bytes": metric_vector(
                [20] * 5, scope=source_scope
            ),
            "logical_read_returned_bytes": metric_vector(
                [20] * 5, scope=source_scope
            ),
            "logical_read_largest_requested_bytes": metric_vector(
                [20] * 5, scope=source_scope
            ),
            "logical_read_largest_returned_bytes": metric_vector(
                [20] * 5, scope=source_scope
            ),
            "logical_read_pattern": pattern_vector(
                ["sequential"] * 5, scope=pattern_scope
            ),
            "compressed_bytes": metric_vector(
                None,
                status="unavailable",
                scope="unavailable_read_at_has_no_compressed_member_boundary",
            ),
            "decompressed_bytes": metric_vector(
                None,
                status="unavailable",
                scope="unavailable_read_at_has_no_decompressed_byte_boundary",
            ),
            "recompressed_bytes": metric_vector(
                None,
                status="unavailable",
                scope="unavailable_atomic_save_has_no_recompressed_byte_boundary",
            ),
            "max_concurrent_reads": metric_vector([1] * 5, scope=source_scope),
        },
        "process": {
            "status": "measured",
            "user_cpu_ticks": metric_vector([1] * 5, scope=process_scope),
            "system_cpu_ticks": metric_vector([1] * 5, scope=process_scope),
            "clock_ticks_per_second": metric_vector([100] * 5, scope=process_scope),
            "minor_faults": metric_vector([1] * 5, scope=process_scope),
            "major_faults": metric_vector(
                [2] * 5,
                scope=process_scope,
            ),
            "voluntary_context_switches": metric_vector([1] * 5, scope=process_scope),
            "nonvoluntary_context_switches": metric_vector(
                [1] * 5, scope=process_scope
            ),
            "rss_delta_bytes": metric_vector(
                [1] * 5, scope="procfs_operation_delta_not_peak"
            ),
            "peak_rss_bytes": metric_vector(
                [1] * 5, scope="process_lifetime_high_water_after_not_operation_peak"
            ),
            "rchar": metric_vector(
                [3] * 5,
                scope=proc_io_scope,
            ),
            "read_bytes": metric_vector(
                [4] * 5,
                scope=proc_io_scope,
            ),
            "wchar": metric_vector([3] * 5, scope=proc_io_scope),
            "write_bytes": metric_vector([3] * 5, scope=proc_io_scope),
            "cancelled_write_bytes": metric_vector([0] * 5, scope=proc_io_scope),
            "syscr": metric_vector(
                [1] * 5,
                scope=proc_io_scope,
            ),
            "syscw": metric_vector([1] * 5, scope=proc_io_scope),
        },
        "sink": {
            "status": "not_applicable",
            "output_bytes": metric_vector(
                None,
                status="not_applicable",
                scope="post_operation_output_length_not_sink_write_volume",
            ),
            "write_status": "measured",
            "accepted_bytes": metric_vector([100] * 5),
            "write_calls": metric_vector([2] * 5),
            "largest_write": metric_vector([64] * 5),
            "write_size_buckets": {
                "status": "measured",
                "bytes_0": metric_vector([0] * 5),
                "bytes_1_to_512": metric_vector([2] * 5),
                "bytes_513_to_4096": metric_vector([0] * 5),
                "bytes_4097_to_16384": metric_vector([0] * 5),
                "bytes_16385_to_65536": metric_vector([0] * 5),
                "bytes_over_65536": metric_vector([0] * 5),
            },
        },
        "publication": {
            "status": "not_applicable",
            "changed_spans": not_applicable("logical_publication_counter"),
            "published_bytes": not_applicable("logical_publication_counter"),
        },
        "materialization": {
            "status": "not_applicable",
            "opc_parts": not_applicable("logical_materialization_counter"),
        },
        "cfb_phases": {
            "status": "not_applicable",
            "open": {
                "elapsed_ns": not_applicable("timed_cfb_phase_elapsed_ns"),
                "logical_read_calls": not_applicable(cfb_source_scope),
                "logical_read_requested_bytes": not_applicable(cfb_source_scope),
                "logical_read_returned_bytes": not_applicable(cfb_source_scope),
            },
            "plan": {
                "elapsed_ns": not_applicable("timed_cfb_phase_elapsed_ns"),
                "logical_read_calls": not_applicable(cfb_source_scope),
                "logical_read_requested_bytes": not_applicable(cfb_source_scope),
                "logical_read_returned_bytes": not_applicable(cfb_source_scope),
            },
            "atomic_publication": {
                "elapsed_ns": not_applicable("timed_cfb_phase_elapsed_ns"),
                "logical_read_calls": not_applicable(cfb_source_scope),
                "logical_read_requested_bytes": not_applicable(cfb_source_scope),
                "logical_read_returned_bytes": not_applicable(cfb_source_scope),
            },
        },
    }


def filesystem_xlsx_operation_metrics_report_fields(sample_indices=None):
    operation_metrics = operation_metrics_report_fields()
    source = operation_metrics["source"]
    source["status"] = "not_applicable"
    source["counter_scope"] = "not_applicable_filesystem_xlsx"
    for field, vector in source.items():
        if field in {"status", "counter_scope"}:
            continue
        vector["status"] = "not_applicable"
        vector.pop("values", None)
    if sample_indices is not None:
        operation_metrics["sample_indices"] = sample_indices
    return operation_metrics


def opc_source_materialize_operation_metrics_report_fields():
    operation_metrics = operation_metrics_report_fields()
    operation_metrics[
        "latency_claim"
    ] = perf_compare.OPC_SOURCE_MATERIALIZATION_EVIDENCE_ONLY_LATENCY_CLAIM
    source = operation_metrics["source"]
    source_scope = perf_compare.OPC_SOURCE_MATERIALIZATION_SOURCE_SCOPE
    source["counter_scope"] = source_scope
    for field in (
        "logical_read_calls",
        "logical_read_returned_bytes",
        "max_concurrent_reads",
    ):
        source[field] = metric_vector([10] * 5, scope=source_scope)
    for field in (
        "logical_read_requested_bytes",
        "logical_read_largest_requested_bytes",
        "logical_read_largest_returned_bytes",
    ):
        scope = (
            perf_compare.OPC_SOURCE_MATERIALIZATION_REQUEST_SCOPE
            if field == "logical_read_requested_bytes"
            else perf_compare.OPC_SOURCE_MATERIALIZATION_LARGEST_SCOPE
        )
        source[field] = metric_vector(
            None,
            status="unavailable",
            scope=scope,
        )
    source["logical_read_pattern"] = pattern_vector(
        None,
        status="unavailable",
        scope="operation_logical_read_at_range_order_not_physical_io",
    )

    sink = operation_metrics["sink"]
    sink["status"] = "not_applicable"
    sink["output_bytes"] = metric_vector(
        None,
        status="not_applicable",
        scope="post_operation_output_length_not_sink_write_volume",
    )
    sink["write_status"] = "not_applicable"
    for field, scope in (
        ("accepted_bytes", "logical_sink_accepted_write_bytes"),
        ("write_calls", "logical_sink_accepted_write_calls"),
        ("largest_write", "logical_sink_largest_accepted_write"),
    ):
        sink[field] = metric_vector(None, status="not_applicable", scope=scope)
    sink["write_size_buckets"] = {
        "status": "not_applicable",
        **{
            field: metric_vector(None, status="not_applicable", scope="logical_sink_accepted_write_size_bucket_counts")
            for field in (
                "bytes_0",
                "bytes_1_to_512",
                "bytes_513_to_4096",
                "bytes_4097_to_16384",
                "bytes_16385_to_65536",
                "bytes_over_65536",
            )
        },
    }
    operation_metrics["materialization"] = {
        "status": "measured",
        "opc_parts": metric_vector([3] * 5, scope="logical_materialization_counter"),
    }
    return operation_metrics


def opc_source_materialize_accounted_operation_metrics_report_fields():
    operation_metrics = opc_source_materialize_operation_metrics_report_fields()
    operation_metrics["latency_claim"] = (
        perf_compare.OPC_MATERIALIZATION_ZIP_LATENCY_CLAIM
    )
    zip_metrics = opc_zip_operation_metrics_report_fields()["opc_zip"]
    for key, value in zip_metrics.items():
        if isinstance(value, dict):
            value["scope"] = perf_compare.OPC_MATERIALIZATION_ZIP_SCOPE
    zip_metrics["scope"] = perf_compare.OPC_MATERIALIZATION_ZIP_SCOPE
    operation_metrics["opc_zip"] = zip_metrics
    return operation_metrics


def opc_source_materialize_report_pair(
    *, accounted=False, comparison_policy=None
):
    case = (
        perf_compare.OPC_SOURCE_MATERIALIZE_ACCOUNTED_CASE
        if accounted
        else perf_compare.OPC_SOURCE_MATERIALIZE_CASE
    )
    oracle = (
        perf_compare.OPC_SOURCE_MATERIALIZE_ORACLE_VERSION if accounted else None
    )
    metrics_factory = (
        opc_source_materialize_accounted_operation_metrics_report_fields
        if accounted
        else opc_source_materialize_operation_metrics_report_fields
    )
    baseline = report()
    current = report(revision="current")
    for item in (baseline, current):
        item["configuration"]["cases"] = [case]
        if oracle is not None:
            item["configuration"][
                perf_compare.OPC_SOURCE_MATERIALIZE_ORACLE_CONFIG_FIELD
            ] = oracle
        item["results"][0]["case"] = case
        item["results"][0]["operation_metrics"] = metrics_factory()
        if accounted:
            item["results"][0]["elapsed_ns"]["sample_order"] = list(range(5))
    selected_policy = (
        policy() if comparison_policy is None else copy.deepcopy(comparison_policy)
    )
    selected_policy["required_cases"] = [case]
    corpus_identity = json.dumps(
        baseline["results"][0]["corpus"],
        sort_keys=True,
        separators=(",", ":"),
        allow_nan=False,
    )
    selected_policy["expected_result_keys_sha256"] = (
        perf_compare.result_key_manifest_sha256([(case, corpus_identity)])
    )
    return baseline, current, selected_policy


def opc_zip_operation_metrics_report_fields(status="measured", values=None):
    operation_metrics = operation_metrics_report_fields()
    operation_metrics["latency_claim"] = (
        perf_compare.OPC_ZIP_EVIDENCE_ONLY_LATENCY_CLAIM
    )
    values = values or {}
    fields = (
        "compressed_deflate_payload_bytes_read",
        "stored_payload_bytes_read",
        "stored_payload_bytes_accepted",
        "deflate_bytes_produced",
        "deflate_bytes_accepted",
        "generated_deflate_payload_bytes_emitted",
        "stored_payload_bytes_emitted",
        "precompressed_payload_bytes_emitted",
        "raw_unchanged_source_bytes_accepted",
        "output_bytes_accepted",
    )
    operation_metrics["opc_zip"] = {
        "status": status,
        "scope": perf_compare.OPC_ZIP_SCOPE,
    }
    for field in fields:
        operation_metrics["opc_zip"][field] = metric_vector(
            values.get(field, [0] * 5) if status == "measured" else None,
            status=status,
            scope=perf_compare.OPC_ZIP_SCOPE,
        )
    return operation_metrics


def add_opc_zip_sample_order(*results):
    for result in results:
        result["elapsed_ns"]["sample_order"] = list(range(5))


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
)


def allocator_policy_fixture():
    comparison_policy = policy()
    comparison_policy["tool_identity"] = {
        **copy.deepcopy(TOOL),
        "binary": "litchi-perf-baseline-alloc",
        "instrumentation": "system_allocator_operation_scoped",
    }
    comparison_policy["expected_build_identity"] = {
        "allocator": "CountingSystemAllocator(std::alloc::System)"
    }
    comparison_policy["expected_result_count"] = 2
    comparison_policy["required_cases"] = ["opc_file_eager_open"]
    comparison_policy["result_key_fields"] = ["case", "corpus", "cache_state"]
    comparison_policy["require_sample_child_identity"] = True
    corpus = report()["results"][0]["corpus"]
    corpus_identity = json.dumps(
        corpus, sort_keys=True, separators=(",", ":"), allow_nan=False
    )
    comparison_policy["expected_result_keys_sha256"] = perf_compare.result_key_manifest_sha256(
        [
            ("opc_file_eager_open", corpus_identity, "warm"),
            ("opc_file_eager_open", corpus_identity, "cold-requested"),
        ]
    )
    comparison_policy["expected_configuration"] = {
        "samples_per_case": 5,
        "filesystem_cache_states": ["warm", "cold-requested"],
        "filesystem_fresh_child_per_sample": True,
        "filesystem_process_isolated": True,
        "filesystem_root_selected": False,
    }
    comparison_policy["metric_classes"] = [
        {
            "name": "allocation",
            "max_regression_percent": 5.0,
            "path_globs": ["operation_metrics/allocation/*/values"],
            "presence": "required",
        }
    ]
    return comparison_policy


def allocator_operation_metrics(value=100, sample_count=5):
    operation_metrics = operation_metrics_report_fields()
    allocation = {
        "status": "measured",
        "scope": "operation_global_system_allocator",
    }
    for field in ALLOCATOR_VECTOR_FIELDS:
        allocation[field] = {
            "values": [value] * sample_count,
            "status": "measured",
            "scope": "operation_global_system_allocator",
        }
    operation_metrics["allocation"] = allocation
    return operation_metrics


def allocator_raw_metrics(value=100):
    return {
        "status": "measured",
        "scope": "operation_global_system_allocator",
        **{field: value for field in ALLOCATOR_VECTOR_FIELDS},
    }


def allocator_report(value=100, revision="baseline"):
    result = report(value=value, revision=revision)
    result["tool"]["binary"] = "litchi-perf-baseline-alloc"
    result["tool"]["instrumentation"] = "system_allocator_operation_scoped"
    result["binary_identity"]["path"] = (
        "/tmp/litchi-perf-alloc-candidate"
        if revision != "baseline"
        else "/tmp/litchi-perf-alloc-baseline"
    )
    result["binary_identity"]["binary_sha256"] = (
        "d" * 64 if revision != "baseline" else "c" * 64
    )
    result["binary_identity"]["binary_bytes"] = 220 if revision != "baseline" else 210
    result["environment"]["allocator"] = "CountingSystemAllocator(std::alloc::System)"
    result["configuration"]["cases"] = ["opc_file_eager_open"]
    result["configuration"]["filesystem_cache_states"] = ["warm", "cold-requested"]
    result["configuration"]["filesystem_fresh_child_per_sample"] = True
    result["configuration"]["filesystem_process_isolated"] = True
    result["configuration"]["filesystem_root_selected"] = False
    result["filesystem_evidence"] = [
        {
            "case": "opc_file_eager_open",
            "sample_count": 5,
            "cache_states": ["warm", "cold-requested"],
            "fresh_child_per_sample": True,
            "samples": [
                {
                    "sample_index": sample_index,
                    "cache_state": cache_state,
                    "child_process_id": (
                        (1_000 if cache_state == "warm" else 2_000) + sample_index
                    ),
                    "allocation_metrics": allocator_raw_metrics(value),
                }
                for cache_state in ("warm", "cold-requested")
                for sample_index in range(5)
            ],
        }
    ]
    first = result["results"][0]
    first["case"] = "opc_file_eager_open"
    first["cache_state"] = "warm"
    first["operation_metrics"] = allocator_operation_metrics(value)
    second = copy.deepcopy(first)
    second["cache_state"] = "cold-requested"
    result["results"].append(second)
    return result


def operation_allocator_policy_fixture():
    comparison_policy = policy()
    comparison_policy["allocator_evidence_scope"] = "operation"
    comparison_policy["tool_identity"] = {
        **copy.deepcopy(TOOL),
        "binary": "litchi-perf-baseline-alloc",
        "instrumentation": "system_allocator_operation_scoped",
    }
    comparison_policy["expected_build_identity"] = {
        "allocator": "CountingSystemAllocator(std::alloc::System)"
    }
    comparison_policy["required_cases"] = ["opc_source_materialize"]
    comparison_policy["expected_result_keys_sha256"] = perf_compare.result_key_manifest_sha256(
        [
            (
                "opc_source_materialize",
                json.dumps(
                    report()["results"][0]["corpus"],
                    sort_keys=True,
                    separators=(",", ":"),
                    allow_nan=False,
                ),
            )
        ]
    )
    comparison_policy["expected_configuration"] = {"samples_per_case": 5}
    comparison_policy["metric_classes"] = [
        {
            "name": "allocation",
            "max_regression_percent": 0.0,
            "path_globs": [
                "operation_metrics/allocation/allocation_calls/values",
                "operation_metrics/allocation/allocated_bytes/values",
            ],
            "presence": "required",
        },
        {
            "name": "invariant_work",
            "max_regression_percent": 0.0,
            "path_globs": [
                "operation_metrics/source/logical_read_calls/values",
                "operation_metrics/source/logical_read_returned_bytes/values",
                "operation_metrics/materialization/opc_parts/values",
            ],
            "presence": "required",
        },
    ]
    return comparison_policy


def operation_allocator_metrics(value=100):
    operation_metrics = opc_source_materialize_operation_metrics_report_fields()
    allocation = {
        "status": "measured",
        "scope": "operation_global_system_allocator",
    }
    for field in ALLOCATOR_VECTOR_FIELDS:
        allocation[field] = {
            "values": [value] * 5,
            "status": "measured",
            "scope": "operation_global_system_allocator",
        }
    operation_metrics["allocation"] = allocation
    return operation_metrics


def operation_allocator_report(value=100, revision="baseline"):
    result = report(value=value, revision=revision)
    result["tool"]["binary"] = "litchi-perf-baseline-alloc"
    result["tool"]["instrumentation"] = "system_allocator_operation_scoped"
    result["binary_identity"]["path"] = (
        "/tmp/litchi-perf-operation-alloc-candidate"
        if revision != "baseline"
        else "/tmp/litchi-perf-operation-alloc-baseline"
    )
    result["binary_identity"]["binary_sha256"] = (
        "d" * 64 if revision != "baseline" else "c" * 64
    )
    result["binary_identity"]["binary_bytes"] = 220 if revision != "baseline" else 210
    result["environment"]["allocator"] = "CountingSystemAllocator(std::alloc::System)"
    result["configuration"]["cases"] = ["opc_source_materialize"]
    first = result["results"][0]
    first["case"] = "opc_source_materialize"
    first["elapsed_ns"]["sample_order"] = list(range(5))
    first["operation_metrics"] = operation_allocator_metrics(value)
    return result


def xlsx_allocator_corpus_fixture():
    corpus = copy.deepcopy(report()["results"][0]["corpus"])
    corpus.update(
        {
            "name": "xlsx-source-repeated-store-medium",
            "generator": "litchi-xlsx-source-repeated-store-corpus-v1",
            "package_format": "XLSX/OPC/ZIP",
            "shape": "medium",
            "archive_sha256": "e" * 64,
        }
    )
    return corpus


def xlsx_allocator_policy_fixture():
    comparison_policy = allocator_policy_fixture()
    comparison_policy["expected_result_count"] = 1
    comparison_policy["required_cases"] = ["xlsx_source_repeated_store_medium"]
    comparison_policy["expected_configuration"]["filesystem_cache_states"] = ["warm"]
    comparison_policy["expected_configuration"]["filesystem_root_selected"] = True
    comparison_policy["filesystem_identity_fields"] = [
        "filesystem_type",
        "source_destination_same_device",
        "storage_identifier",
    ]
    corpus = xlsx_allocator_corpus_fixture()
    corpus_identity = json.dumps(
        corpus, sort_keys=True, separators=(",", ":"), allow_nan=False
    )
    comparison_policy["expected_result_keys_sha256"] = (
        perf_compare.result_key_manifest_sha256(
            [("xlsx_source_repeated_store_medium", corpus_identity, "warm")]
        )
    )
    return comparison_policy, corpus


def xlsx_allocator_report(corpus, value=100, revision="baseline"):
    result = allocator_report(value=value, revision=revision)
    result["configuration"]["cases"] = ["xlsx_source_repeated_store_medium"]
    result["configuration"]["filesystem_cache_states"] = ["warm"]
    result["configuration"]["filesystem_root_selected"] = True
    result["environment"].update(
        filesystem_type="tmpfs",
        source_destination_same_device=True,
        storage_identifier=None,
    )
    row = copy.deepcopy(result["results"][0])
    row["case"] = "xlsx_source_repeated_store_medium"
    row["cache_state"] = "warm"
    row["corpus"] = copy.deepcopy(corpus)
    result["results"] = [row]
    evidence = copy.deepcopy(result["filesystem_evidence"][0])
    evidence["case"] = "xlsx_source_repeated_store_medium"
    evidence["cache_states"] = ["warm"]
    evidence["corpus"] = copy.deepcopy(corpus)
    evidence["samples"] = evidence["samples"][:5]
    for sample in evidence["samples"]:
        sample["xlsx_source_sha256"] = corpus["archive_sha256"]
    result["filesystem_evidence"] = [evidence]
    return result


def descriptive_parallel_report(value=100, revision="baseline"):
    measured = report(value, revision)
    result = measured["results"][0]
    result["elapsed_ns"].update(
        samples=[100, 101, 102, 103, 104],
        min=100,
        p50=102,
        p95=104,
        p99=104,
        max=104,
    )
    result["elapsed_ns"]["sample_order"] = [1, 0, 2, 4, 3]
    result["execution"] = {"worker_count": 2, "logical_tasks": 5}
    result["source"] = {
        "read_calls": [10] * 5,
        "simulation": {"physical_request_count": [10, 20, 30, 40, 50]},
    }
    measured["parallel_metrics"] = {
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
        "cases": [
            {
                "case": "opc_open",
                "corpus_sha256": "abc",
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
                    "value": 5,
                    "scope": "result.execution.logical_tasks",
                },
                "deterministic_chunk_count": {
                    "status": "measured",
                    "value": [20, 10, 30, 50, 40],
                    "scope": "result.source.simulation.physical_request_count",
                },
                "lock_wait_ns": {
                    "status": "unavailable",
                    "scope": "lock_wait_ns",
                    "reason": "no exact instrumented lock boundary is present",
                },
            }
        ],
    }
    return measured


def descriptive_parallel_lock_report(revision="baseline"):
    measured = descriptive_parallel_report(revision=revision)
    configuration = measured["configuration"]
    configuration["opc_cache_lock_diagnostics"] = True
    source = measured["results"][0]["source"]
    source["opc_cache"] = {
        "worker_count": 2,
        "persistent_worker_teams_created": 1,
        "lock_diagnostics": {
            "scope": (
                "opc_cache_direct_mutex_lock_acquisition_including_"
                "observer_timer_overhead"
            ),
            "excluded": "same_part_condvar_wait_timeout_and_mutex_reacquisition",
            "coverage": (
                "all_worker_part_data_requests_including_pre_admission_"
                "rendezvous"
            ),
            "cache_lock_acquisitions": [1, 2, 1, 2, 1],
            "cache_lock_wait_ns": [2, 1, 2, 3, 5],
            "flight_lock_acquisitions": [2, 1, 3, 1, 2],
            "flight_lock_wait_ns": [7, 2, 3, 4, 6],
            "total_lock_acquisitions": [3, 3, 4, 3, 3],
            "total_lock_wait_ns": [9, 3, 5, 7, 11],
        },
    }
    case = measured["parallel_metrics"]["cases"][0]
    case["observed_local_worker_count"] = {
        "status": "measured",
        "value": 2,
        "scope": (
            "result.source.opc_cache.worker_count_with_one_"
            "created_local_worker_team"
        ),
    }
    case["lock_wait_ns"] = {
        "status": "measured",
        "value": 7,
        "scope": (
            "result.source.opc_cache.lock_diagnostics."
            "total_lock_wait_ns.p50"
        ),
    }
    return measured


class PerfCompareTests(unittest.TestCase):
    def operation_metrics_policy(self):
        comparison_policy = policy()
        comparison_policy["metric_classes"][-1]["path_globs"].extend(
            [
                "**/*faults",
                "**/*read_bytes",
                "**/*write_calls",
                "**/*logical_read_requested_bytes",
                "**/*logical_read_returned_bytes",
                "**/*logical_read_largest_requested_bytes",
                "**/*logical_read_largest_returned_bytes",
                "**/*max_concurrent_reads",
            ]
        )
        return comparison_policy

    def test_checked_policy_pins_identity_only_default_manifest(self):
        repository = Path(__file__).resolve().parents[1]
        checked_policy = json.loads(
            (repository / "docs/performance/perf-regression-policy-v1.json").read_text(
                encoding="utf-8"
            )
        )
        manifest = json.loads(
            (
                repository
                / "docs/performance/results/perf-regression-default-manifest-v1.json"
            ).read_text(encoding="utf-8")
        )

        self.assertEqual(manifest["manifest_kind"], "case-corpus-key-identity")
        self.assertEqual(manifest["result_count"], 198)
        self.assertEqual(manifest["case_count"], 36)
        self.assertEqual(manifest["source_report_samples_per_case"], 1)
        self.assertEqual(manifest["source_report_warmup_iterations_per_case"], 0)
        self.assertEqual(checked_policy["schema_version"], 2)
        self.assertEqual(checked_policy["tool_identity"]["binary"], "litchi-perf-baseline")
        self.assertEqual(checked_policy["tool_identity"]["instrumentation"], "none")
        allocation_class = next(
            item
            for item in checked_policy["metric_classes"]
            if item["name"] == "allocation"
        )
        self.assertIn("**/*allocation_calls/values", allocation_class["path_globs"])
        self.assertIn("**/*peak_live_bytes_after/values", allocation_class["path_globs"])
        self.assertEqual(manifest["default_cases"], checked_policy["required_cases"])
        self.assertEqual(
            set(checked_policy["latency_thresholds_percent"]),
            {"p50", "p95", "p99"},
        )
        self.assertTrue(
            all(
                metric_class["presence"] == "optional"
                for metric_class in checked_policy["metric_classes"]
            )
        )
        for field, expected in manifest["identity_configuration"].items():
            self.assertEqual(
                checked_policy["expected_configuration"][field], expected, field
            )

        keys = []
        case_corpora = manifest["case_corpora"]
        corpora = manifest["corpora"]
        self.assertEqual(set(case_corpora), set(checked_policy["required_cases"]))
        for case in manifest["default_cases"]:
            names = case_corpora[case]
            self.assertEqual(len(names), len(set(names)), case)
            for name in names:
                corpus = corpora[name]
                corpus_identity = json.dumps(
                    corpus, sort_keys=True, separators=(",", ":"), allow_nan=False
                )
                keys.append((case, corpus_identity))

        self.assertEqual(len(keys), manifest["result_count"])
        self.assertEqual(len(keys), len(set(keys)))
        self.assertEqual(
            set(corpora), {name for names in case_corpora.values() for name in names}
        )
        self.assertEqual(
            sorted(len(names) for names in case_corpora.values()),
            [3] * 18 + [8] * 18,
        )
        for forbidden in ("elapsed_ns", "metrics", "sink", "output_sha256"):
            self.assertNotIn(forbidden, manifest)
        digest = perf_compare.result_key_manifest_sha256(keys)
        self.assertEqual(digest, manifest["result_keys_sha256"])
        self.assertEqual(digest, checked_policy["expected_result_keys_sha256"])

    def test_additive_corpus_catalog_reference_preserves_v1_comparison(self):
        baseline = report()
        current = report(revision="current")
        reference = {
            "manifest_version": 2,
            "catalog_id": "litchi-perf-corpus-v2",
            "catalog_sha256": "0" * 64,
            "content_set_sha256": "1" * 64,
        }
        baseline["corpus_catalog"] = copy.deepcopy(reference)
        current["corpus_catalog"] = copy.deepcopy(reference)
        comparison = perf_compare.compare_reports(baseline, current, policy())
        self.assertEqual(comparison["status"], "pass")
        self.assertEqual(comparison["summary"]["matched_results"], 1)

    def test_schema_one_report_without_corpus_catalog_remains_supported(self):
        baseline = report()
        current = report(revision="current")
        comparison = perf_compare.compare_reports(baseline, current, policy())
        self.assertEqual(comparison["status"], "pass")

    def test_schema_one_descriptorless_reports_require_explicit_legacy_policy_opt_out(self):
        legacy_policy = policy()
        legacy_policy.pop("require_binary_identity")
        baseline = report()
        current = report(revision="current")
        baseline.pop("binary_identity")
        current.pop("binary_identity")
        comparison = perf_compare.compare_reports(baseline, current, legacy_policy)
        self.assertEqual(comparison["status"], "pass")

    def test_binary_identity_is_required_validated_and_not_equal_across_reports(self):
        baseline = report()
        current = report(revision="current")
        comparison = perf_compare.compare_reports(baseline, current, policy())
        self.assertEqual(comparison["status"], "pass")
        self.assertNotEqual(
            comparison["baseline_binary_sha256"], comparison["current_binary_sha256"]
        )

        missing = copy.deepcopy(current)
        missing.pop("binary_identity")
        with self.assertRaisesRegex(
            perf_compare.ComparisonInputError, "binary_identity"
        ):
            perf_compare.compare_reports(baseline, missing, policy())

        malformed = copy.deepcopy(current)
        malformed["binary_identity"]["binary_sha256"] = "z" * 64
        with self.assertRaisesRegex(
            perf_compare.ComparisonInputError, "binary_sha256"
        ):
            perf_compare.compare_reports(baseline, malformed, policy())

        non_executable_mode = copy.deepcopy(current)
        non_executable_mode["binary_identity"]["mode_bits"] = 0o644
        with self.assertRaisesRegex(
            perf_compare.ComparisonInputError, "executable permission bits"
        ):
            perf_compare.compare_reports(baseline, non_executable_mode, policy())

        missing_unix_mode = copy.deepcopy(current)
        missing_unix_mode["binary_identity"]["mode_bits"] = None
        with self.assertRaisesRegex(
            perf_compare.ComparisonInputError, "present for Unix targets"
        ):
            perf_compare.compare_reports(baseline, missing_unix_mode, policy())

        oversized = copy.deepcopy(current)
        oversized["binary_identity"]["binary_bytes"] = 1 << 64
        with self.assertRaisesRegex(
            perf_compare.ComparisonInputError, "positive unsigned integer"
        ):
            perf_compare.compare_reports(baseline, oversized, policy())

        drift = copy.deepcopy(current)
        drift["binary_identity"]["mode_bits"] = 0o700
        # Baseline/current binaries are allowed to differ in every descriptor
        # identity field, provided each descriptor remains syntactically valid.
        self.assertEqual(
            perf_compare.compare_reports(baseline, drift, policy())["status"], "pass"
        )

    def test_allocator_policy_accepts_only_allocator_binary_identity(self):
        repository = Path(__file__).resolve().parents[1]
        allocator_policy = json.loads(
            (
                repository
                / "docs/performance/perf-regression-policy-allocator-v1.json"
            ).read_text(encoding="utf-8")
        )
        perf_compare.validate_policy(allocator_policy)
        self.assertEqual(
            allocator_policy["tool_identity"]["binary"],
            "litchi-perf-baseline-alloc",
        )
        self.assertEqual(
            allocator_policy["tool_identity"]["instrumentation"],
            "system_allocator_operation_scoped",
        )
        self.assertEqual(
            allocator_policy["expected_build_identity"]["allocator"],
            "CountingSystemAllocator(std::alloc::System)",
        )
        allocator_manifest = json.loads(
            (
                repository
                / "docs/performance/results/perf-regression-allocator-manifest-v1.json"
            ).read_text(encoding="utf-8")
        )
        self.assertEqual(
            allocator_policy["expected_result_keys_sha256"],
            allocator_manifest["result_keys_sha256"],
        )
        corpus = allocator_manifest["corpora"]["few-large-incompressible"]
        corpus_identity = json.dumps(
            corpus, sort_keys=True, separators=(",", ":"), allow_nan=False
        )
        self.assertEqual(
            allocator_manifest["result_keys_sha256"],
            perf_compare.result_key_manifest_sha256(
                [
                    ("opc_file_eager_open", corpus_identity, "warm"),
                    ("opc_file_eager_open", corpus_identity, "cold-requested"),
                ]
            ),
        )
        self.assertEqual(
            allocator_policy["result_key_fields"], ["case", "corpus", "cache_state"]
        )
        self.assertEqual(
            [item["name"] for item in allocator_policy["metric_classes"]],
            ["allocation"],
        )
        self.assertEqual(allocator_policy["metric_classes"][0]["presence"], "required")
        invalid_identity = copy.deepcopy(allocator_policy)
        invalid_identity["tool_identity"]["binary"] = "litchi-perf-baseline"
        with self.assertRaisesRegex(
            perf_compare.ComparisonInputError, "allocator policy.*binary"
        ):
            perf_compare.validate_policy(invalid_identity)
        allocator_report = report()
        allocator_report["tool"]["binary"] = "litchi-perf-baseline-alloc"
        allocator_report["tool"]["instrumentation"] = "system_allocator_operation_scoped"
        with self.assertRaisesRegex(perf_compare.ComparisonInputError, "tool does not match"):
            perf_compare.compare_reports(allocator_report, report(revision="current"), policy())

    def test_operation_allocator_policy_compares_only_allocation_and_invariant_work(self):
        comparison_policy = operation_allocator_policy_fixture()
        baseline = operation_allocator_report()
        current = operation_allocator_report(revision="current")
        result = perf_compare.compare_reports(baseline, current, comparison_policy)
        self.assertEqual(result["status"], "pass")
        self.assertEqual(result["summary"]["matched_results"], 1)
        self.assertEqual(result["summary"]["compared_metrics"], 5)
        self.assertEqual(result["summary"]["latency_claims"], "withheld_instrumentation")
        self.assertEqual(result["summary"]["latency_compared_results"], 0)
        self.assertEqual(result["summary"]["latency_excluded_results"], 1)
        self.assertEqual(
            {item["metric"] for item in result["comparisons"]},
            {
                "operation_metrics.allocation.allocation_calls.values",
                "operation_metrics.allocation.allocated_bytes.values",
                "operation_metrics.materialization.opc_parts.values",
                "operation_metrics.source.logical_read_calls.values",
                "operation_metrics.source.logical_read_returned_bytes.values",
            },
        )
        self.assertTrue(all("cache_state" not in item for item in result["comparisons"]))

    def test_checked_0387_operation_allocator_policy_and_comparison_are_scoped(self):
        repository = Path(__file__).resolve().parents[1]
        raw_artifacts = repository / "docs/performance/results/change-0387"
        checked_policy = json.loads(
            (
                repository
                / "docs/performance/perf-regression-policy-opc-source-materialize-allocator-v1.json"
            ).read_text(encoding="utf-8")
        )
        checked_comparison = json.loads(
            (
                repository
                / "docs/performance/results/opc-source-materialization-shared-0388-comparison.json"
            ).read_text(encoding="utf-8")
        )
        perf_compare.validate_policy(checked_policy)
        self.assertEqual(checked_policy["allocator_evidence_scope"], "operation")
        self.assertEqual(checked_policy["result_key_fields"], ["case", "corpus"])
        self.assertEqual(checked_comparison["status"], "pass")
        self.assertEqual(
            checked_comparison["policy"]["policy_id"], checked_policy["policy_id"]
        )
        self.assertEqual(checked_comparison["summary"]["matched_results"], 3)
        self.assertEqual(checked_comparison["summary"]["compared_metrics"], 15)
        self.assertEqual(
            checked_comparison["summary"]["latency_claims"],
            "withheld_instrumentation",
        )
        self.assertEqual(checked_comparison["summary"]["latency_compared_results"], 0)
        self.assertEqual(checked_comparison["summary"]["latency_excluded_results"], 3)
        self.assertTrue(
            all("cache_state" not in item for item in checked_comparison["comparisons"])
        )
        self.assertEqual(
            {item["metric"] for item in checked_comparison["comparisons"]},
            {
                "operation_metrics.allocation.allocation_calls.values",
                "operation_metrics.allocation.allocated_bytes.values",
                "operation_metrics.materialization.opc_parts.values",
                "operation_metrics.source.logical_read_calls.values",
                "operation_metrics.source.logical_read_returned_bytes.values",
            },
        )

        def read_zstd_json(path):
            try:
                payload = subprocess.check_output(
                    ["zstd", "-q", "-d", "-c", str(path)]
                )
            except FileNotFoundError:
                self.skipTest("zstd is unavailable; cannot rederive checked artifacts")
            return json.loads(payload)

        artifact_pairs = (
            ("tiny", "tiny-compressible", "compressible"),
            ("many-small", "many-small-incompressible", "incompressible"),
            ("few-large", "few-large-incompressible", "incompressible"),
        )

        def aggregate(role):
            report_names = [
                f"{role}-{filename}.json.zst"
                for _, filename, _ in artifact_pairs
            ]
            reports = [read_zstd_json(raw_artifacts / name) for name in report_names]
            aggregate_report = copy.deepcopy(reports[0])
            aggregate_report["results"] = [
                report_value["results"][0] for report_value in reports
            ]
            aggregate_report.pop("parallel_metrics", None)
            aggregate_report.pop("corpus_catalog", None)
            aggregate_report["configuration"]["corpus_shapes"] = [
                shape for shape, _, _ in artifact_pairs
            ]
            aggregate_report["configuration"]["payload_kinds"] = [
                payload_kind
                for payload_kind in dict.fromkeys(
                    payload_kind for _, _, payload_kind in artifact_pairs
                )
            ]
            return aggregate_report

        self.assertEqual(
            perf_compare.compare_reports(
                aggregate("control"), aggregate("candidate"), checked_policy
            ),
            checked_comparison,
        )

    def test_checked_0390_decoder_session_allocator_evidence_rederives_in_memory(self):
        repository = Path(__file__).resolve().parents[1]
        raw_artifacts = repository / "docs/performance/results/change-0390"
        checked_policy = json.loads(
            (
                repository
                / "docs/performance/perf-regression-policy-opc-source-materialize-allocator-v1.json"
            ).read_text(encoding="utf-8")
        )
        checked_comparison = json.loads(
            (
                repository
                / "docs/performance/results/opc-source-materialization-decoder-session-0390-comparison.json"
            ).read_text(encoding="utf-8")
        )
        perf_compare.validate_policy(checked_policy)
        self.assertEqual(checked_policy["allocator_evidence_scope"], "operation")
        self.assertEqual(checked_policy["result_key_fields"], ["case", "corpus"])
        self.assertEqual(checked_comparison["status"], "pass")
        self.assertEqual(
            checked_comparison["policy"]["policy_id"], checked_policy["policy_id"]
        )
        self.assertEqual(checked_comparison["summary"]["matched_results"], 3)
        self.assertEqual(checked_comparison["summary"]["compared_metrics"], 15)
        self.assertEqual(
            checked_comparison["summary"]["latency_claims"],
            "withheld_instrumentation",
        )
        self.assertEqual(checked_comparison["summary"]["latency_compared_results"], 0)
        self.assertEqual(checked_comparison["summary"]["latency_excluded_results"], 3)
        self.assertTrue(
            all("cache_state" not in item for item in checked_comparison["comparisons"])
        )
        self.assertEqual(
            {item["metric"] for item in checked_comparison["comparisons"]},
            {
                "operation_metrics.allocation.allocation_calls.values",
                "operation_metrics.allocation.allocated_bytes.values",
                "operation_metrics.materialization.opc_parts.values",
                "operation_metrics.source.logical_read_calls.values",
                "operation_metrics.source.logical_read_returned_bytes.values",
            },
        )

        def read_zstd_json(path):
            try:
                payload = subprocess.check_output(
                    ["zstd", "-q", "-d", "-c", str(path)]
                )
            except FileNotFoundError:
                self.skipTest("zstd is unavailable; cannot rederive checked artifacts")
            return json.loads(payload)

        artifact_pairs = (
            ("tiny", "tiny-compressible", "compressible"),
            ("many-small", "many-small-incompressible", "incompressible"),
            ("few-large", "few-large-incompressible", "incompressible"),
        )

        def aggregate(role):
            report_names = [
                f"{role}-{filename}.json.zst"
                for _, filename, _ in artifact_pairs
            ]
            reports = [read_zstd_json(raw_artifacts / name) for name in report_names]
            aggregate_report = copy.deepcopy(reports[0])
            aggregate_report["results"] = [
                report_value["results"][0] for report_value in reports
            ]
            aggregate_report.pop("parallel_metrics", None)
            aggregate_report.pop("corpus_catalog", None)
            aggregate_report["configuration"]["corpus_shapes"] = [
                shape for shape, _, _ in artifact_pairs
            ]
            aggregate_report["configuration"]["payload_kinds"] = [
                payload_kind
                for payload_kind in dict.fromkeys(
                    payload_kind for _, _, payload_kind in artifact_pairs
                )
            ]
            return aggregate_report

        self.assertEqual(
            perf_compare.compare_reports(
                aggregate("control"), aggregate("candidate"), checked_policy
            ),
            checked_comparison,
        )

    def test_operation_allocator_policy_rejects_malformed_envelope(self):
        comparison_policy = operation_allocator_policy_fixture()
        baseline = operation_allocator_report()
        mutations = (
            (
                lambda report: report["results"][0]["operation_metrics"]["allocation"].__setitem__(
                    "status", "unavailable"
                ),
                "allocation.status must be 'measured'",
            ),
            (
                lambda report: report["results"][0]["operation_metrics"]["allocation"].__setitem__(
                    "scope", "wrong_scope"
                ),
                "allocation.scope must be",
            ),
            (
                lambda report: report["results"][0]["operation_metrics"]["allocation"][
                    "allocated_bytes"
                ]["values"].pop(),
                "allocated_bytes.values cardinality",
            ),
            (
                lambda report: report["results"][0]["operation_metrics"][
                    "sample_indices"
                ].__setitem__(0, 1),
                "sample_indices must be a permutation",
            ),
            (
                lambda report: report["results"][0]["elapsed_ns"].__setitem__(
                    "sample_order", list(reversed(range(5)))
                ),
                "sample_indices must match elapsed_ns.sample_order",
            ),
        )
        for mutate, message in mutations:
            with self.subTest(message=message):
                current = operation_allocator_report(revision="current")
                mutate(current)
                with self.assertRaisesRegex(
                    perf_compare.ComparisonInputError, message
                ):
                    perf_compare.compare_reports(baseline, current, comparison_policy)

    def test_operation_allocator_policy_rejects_filesystem_assumptions_and_cross_mode_reports(self):
        operation_policy = operation_allocator_policy_fixture()
        filesystem_policy = allocator_policy_fixture()

        invalid_operation_policy = copy.deepcopy(operation_policy)
        invalid_operation_policy["result_key_fields"] = [
            "case",
            "corpus",
            "cache_state",
        ]
        with self.assertRaisesRegex(
            perf_compare.ComparisonInputError, "case/corpus result keys"
        ):
            perf_compare.validate_policy(invalid_operation_policy)

        invalid_operation_policy = copy.deepcopy(operation_policy)
        invalid_operation_policy["expected_configuration"]["filesystem_cache_states"] = [
            "warm"
        ]
        with self.assertRaisesRegex(
            perf_compare.ComparisonInputError, "filesystem configuration fields"
        ):
            perf_compare.validate_policy(invalid_operation_policy)

        operation_baseline = operation_allocator_report()
        operation_current = operation_allocator_report(revision="current")
        operation_current["filesystem_evidence"] = []
        with self.assertRaisesRegex(
            perf_compare.ComparisonInputError, "filesystem_evidence.*forbidden"
        ):
            perf_compare.compare_reports(
                operation_baseline, operation_current, operation_policy
            )

        operation_current = operation_allocator_report(revision="current")
        operation_current["results"][0]["cache_state"] = "warm"
        with self.assertRaisesRegex(
            perf_compare.ComparisonInputError, "cache_state.*forbidden"
        ):
            perf_compare.compare_reports(
                operation_baseline, operation_current, operation_policy
            )

        filesystem_baseline = allocator_report()
        filesystem_current = allocator_report(revision="current")
        for report_value in (filesystem_baseline, filesystem_current):
            report_value["configuration"]["cases"] = ["opc_source_materialize"]
        with self.assertRaisesRegex(
            perf_compare.ComparisonInputError, "filesystem_evidence.*forbidden"
        ):
            perf_compare.compare_reports(
                filesystem_baseline, filesystem_current, operation_policy
            )
        strict_filesystem_baseline = allocator_report()
        filesystem_missing_cache = allocator_report(revision="current")
        filesystem_missing_cache["results"][0].pop("cache_state")
        with self.assertRaisesRegex(
            perf_compare.ComparisonInputError, "cache_state must be a non-empty string"
        ):
            perf_compare.compare_reports(
                strict_filesystem_baseline,
                filesystem_missing_cache,
                filesystem_policy,
            )
        with self.assertRaisesRegex(
            perf_compare.ComparisonInputError, "filesystem_evidence is required"
        ):
            for report_value in (operation_baseline, operation_current):
                report_value["configuration"]["cases"] = ["opc_file_eager_open"]
                report_value["configuration"].update(
                    filesystem_cache_states=["warm", "cold-requested"],
                    filesystem_fresh_child_per_sample=True,
                    filesystem_process_isolated=True,
                    filesystem_root_selected=False,
                )
            perf_compare.compare_reports(
                operation_baseline, operation_current, filesystem_policy
            )

    def test_allocator_policies_require_paired_instrumented_binary_identity(self):
        operation_policy = operation_allocator_policy_fixture()
        missing_identity_requirement = copy.deepcopy(operation_policy)
        missing_identity_requirement.pop("require_binary_identity")
        with self.assertRaisesRegex(
            perf_compare.ComparisonInputError, "must require binary identity"
        ):
            perf_compare.validate_policy(missing_identity_requirement)

        null_scope = copy.deepcopy(operation_policy)
        null_scope["allocator_evidence_scope"] = None
        with self.assertRaisesRegex(
            perf_compare.ComparisonInputError,
            "allocator_evidence_scope must be 'filesystem' or 'operation'",
        ):
            perf_compare.validate_policy(null_scope)

        baseline = operation_allocator_report()
        current = operation_allocator_report(revision="current")
        current.pop("binary_identity")
        with self.assertRaisesRegex(
            perf_compare.ComparisonInputError, "binary_identity"
        ):
            perf_compare.compare_reports(baseline, current, operation_policy)

        ordinary = report(revision="current")
        with self.assertRaisesRegex(
            perf_compare.ComparisonInputError, "tool does not match"
        ):
            perf_compare.compare_reports(baseline, ordinary, operation_policy)

    def test_xlsx_allocator_policy_pins_0269_warm_corpus_keys(self):
        repository = Path(__file__).resolve().parents[1]
        allocator_policy = json.loads(
            (
                repository
                / "docs/performance/perf-regression-policy-xlsx-allocator-v1.json"
            ).read_text(encoding="utf-8")
        )
        allocator_manifest = json.loads(
            (
                repository
                / "docs/performance/results/perf-regression-xlsx-allocator-manifest-v1.json"
            ).read_text(encoding="utf-8")
        )

        evidence_manifest_path = repository / allocator_manifest["evidence_manifest"]
        evidence_manifest = json.loads(evidence_manifest_path.read_text(encoding="utf-8"))
        self.assertEqual(
            hashlib.sha256(evidence_manifest_path.read_bytes()).hexdigest(),
            allocator_manifest["evidence_manifest_sha256"],
        )
        self.assertEqual(evidence_manifest["manifest_kind"], "litchi-perf-abba-artifacts")
        self.assertEqual(evidence_manifest["schema_version"], 1)
        self.assertEqual(evidence_manifest["change_id"], "0269-xlsx-repeated-store-cache-abba")
        self.assertEqual(
            allocator_manifest["harness_source"],
            "tools/perf-baseline/src/filesystem.rs:XlsxRepeatStoreScenario",
        )
        harness_path, harness_symbol = allocator_manifest["harness_source"].split(":", 1)
        harness_source = (repository / harness_path).read_text(encoding="utf-8")
        self.assertIn(harness_symbol, harness_source)
        for case in (
            "xlsx_source_repeated_store_medium",
            "xlsx_source_repeated_store_oversized",
        ):
            self.assertIn(case, harness_source)

        summary_path = evidence_manifest_path.parent / evidence_manifest["summary"]["path"]
        summary = json.loads(summary_path.read_text(encoding="utf-8"))
        self.assertEqual(summary["schema_version"], evidence_manifest["summary"]["schema_version"])
        summary_canonical = perf_abba_package.canonical_json(summary)
        self.assertEqual(
            hashlib.sha256(summary_canonical).hexdigest(),
            evidence_manifest["summary"]["canonical_sha256"],
        )
        self.assertEqual(
            len(summary_canonical), evidence_manifest["summary"]["canonical_bytes"]
        )

        evidence_reports = {}
        for artifact in evidence_manifest["artifacts"]:
            artifact_path = evidence_manifest_path.parent / artifact["path"]
            compressed = artifact_path.read_bytes()
            self.assertEqual(len(compressed), artifact["bytes"])
            self.assertEqual(hashlib.sha256(compressed).hexdigest(), artifact["sha256"])
            raw = subprocess.check_output(["zstd", "-dc", str(artifact_path)])
            self.assertEqual(len(raw), artifact["uncompressed_bytes"])
            self.assertEqual(hashlib.sha256(raw).hexdigest(), artifact["uncompressed_sha256"])
            report = json.loads(raw)
            self.assertEqual(report["schema_version"], allocator_manifest["report_schema_version"])
            canonical = perf_abba_package.canonical_json(report)
            self.assertEqual(
                hashlib.sha256(canonical).hexdigest(), artifact["canonical_sha256"]
            )
            self.assertEqual(
                artifact["canonical_sha256"],
                evidence_manifest["summary"]["report_identity"][artifact["role"]][
                    "canonical_sha256"
                ],
            )
            evidence_reports[artifact["role"]] = report

        self.assertEqual(set(evidence_reports), {"a1", "a2", "b1", "b2"})
        source_cases = {
            "xlsx_source_repeated_store_medium",
            "xlsx_source_repeated_store_oversized",
        }
        source_corpora = {
            result["case"]: result["corpus"]
            for result in evidence_reports["a1"]["results"]
            if result["case"] in source_cases
        }
        self.assertEqual(set(source_corpora), source_cases)
        manifest_corpora = allocator_manifest["corpora"]
        for role, report in evidence_reports.items():
            decompressed_corpora = {
                result["corpus"]["name"]: result["corpus"]
                for result in report["results"]
                if result["case"] in source_cases
            }
            self.assertEqual(
                decompressed_corpora,
                manifest_corpora,
                f"{role} decompressed source corpora differ from allocator manifest",
            )
        for role, report in evidence_reports.items():
            configuration = report["configuration"]
            self.assertEqual(report["tool"]["instrumentation"], "none")
            self.assertEqual(configuration["samples_per_case"], 500)
            self.assertEqual(configuration["warmup_iterations_per_case"], 20)
            self.assertEqual(configuration["filesystem_cache_states"], ["warm"])
            self.assertTrue(configuration["filesystem_fresh_child_per_sample"])
            self.assertTrue(configuration["filesystem_process_isolated"])
            self.assertTrue(configuration["filesystem_root_selected"])
            self.assertEqual(configuration["cases"], sorted(source_cases))
            for result in report["results"]:
                if result["case"] not in source_cases:
                    continue
                self.assertEqual(result["cache_state"], "warm")
                self.assertEqual(result["corpus"], source_corpora[result["case"]])
            for evidence in report["filesystem_evidence"]:
                if evidence["case"] not in source_cases:
                    continue
                self.assertEqual(evidence["corpus"], source_corpora[evidence["case"]])
                self.assertEqual(evidence["sample_count"], 500)
                self.assertEqual(evidence["warmup_iterations"], 20)
                self.assertEqual(evidence["cache_states"], ["warm"])
                self.assertTrue(evidence["fresh_child_per_sample"])
                child_process_ids = [sample["child_process_id"] for sample in evidence["samples"]]
                self.assertEqual(len(child_process_ids), len(set(child_process_ids)))
                for sample in evidence["samples"]:
                    self.assertEqual(
                        sample["xlsx_source_sha256"],
                        source_corpora[evidence["case"]]["archive_sha256"],
                    )

        perf_compare.validate_policy(allocator_policy)
        expected_cases = [
            "xlsx_source_repeated_store_medium",
            "xlsx_source_repeated_store_oversized",
        ]
        self.assertEqual(allocator_policy["required_cases"], expected_cases)
        self.assertEqual(allocator_policy["result_key_fields"], ["case", "corpus", "cache_state"])
        self.assertEqual(
            allocator_policy["tool_identity"]["instrumentation"],
            "system_allocator_operation_scoped",
        )
        self.assertEqual(
            allocator_policy["expected_configuration"],
            {
                "samples_per_case": 30,
                "warmup_iterations_per_case": 3,
                "filesystem_cache_states": ["warm"],
                "filesystem_fresh_child_per_sample": True,
                "filesystem_process_isolated": True,
                "filesystem_root_selected": True,
            },
        )
        self.assertEqual(allocator_manifest["source_report_samples_per_case"], 500)
        self.assertEqual(allocator_manifest["source_report_warmup_iterations_per_case"], 20)
        self.assertEqual(
            allocator_manifest["planned_allocator_policy"],
            {
                "samples_per_case": 30,
                "warmup_iterations_per_case": 3,
                "minimum_samples": 30,
            },
        )
        self.assertEqual(allocator_manifest["manifest_kind"], "allocator-filesystem-case-corpus-cache-key-identity")
        self.assertEqual(allocator_manifest["required_cases"], expected_cases)
        self.assertEqual(allocator_manifest["cache_states"], ["warm"])
        self.assertEqual(allocator_manifest["result_count"], 2)
        self.assertEqual(
            allocator_manifest["case_corpora"],
            {
                "xlsx_source_repeated_store_medium": [
                    "xlsx-source-repeated-store-medium"
                ],
                "xlsx_source_repeated_store_oversized": [
                    "xlsx-source-repeated-store-oversized"
                ],
            },
        )
        expected_corpora = {
            "xlsx-source-repeated-store-medium": {
                "archive_bytes": 4226429,
                "archive_member_count": 17,
                "archive_sha256": "dfff7ec0c749d9e404091776f15a8fb690985af7f58efdfe659dbeaed7145036",
                "generator": "litchi-xlsx-source-repeated-store-corpus-v1",
                "name": "xlsx-source-repeated-store-medium",
                "shape": "medium",
            },
            "xlsx-source-repeated-store-oversized": {
                "archive_bytes": 4236114,
                "archive_member_count": 17,
                "archive_sha256": "3cf797e44ef51189a4b62d040cf39ff2af670ebd909c6e806f387b51e72ecfec",
                "generator": "litchi-xlsx-source-repeated-store-corpus-v1",
                "name": "xlsx-source-repeated-store-oversized",
                "shape": "oversized",
            },
        }
        for name, expected in expected_corpora.items():
            corpus = allocator_manifest["corpora"][name]
            for field, value in expected.items():
                self.assertEqual(corpus[field], value, f"{name}.{field}")
            self.assertEqual(corpus["target_entry"], "Sheet1!A1")
            self.assertEqual(corpus["target_payload_bytes"], 1)
            self.assertEqual(
                corpus["target_payload_sha256"],
                "5feceb66ffc86f38d952786c6d696c79c2dbc239dd4e91b46729d73a27fb57e9",
            )

        keys = []
        for case in expected_cases:
            for name in allocator_manifest["case_corpora"][case]:
                corpus = allocator_manifest["corpora"][name]
                corpus_identity = json.dumps(
                    corpus, sort_keys=True, separators=(",", ":"), allow_nan=False
                )
                keys.append((case, corpus_identity, "warm"))
        digest = perf_compare.result_key_manifest_sha256(keys)
        self.assertEqual(digest, "679776540ae864a066360d84b3e84c4163faff53ef8398891179892080b8a86e")
        self.assertEqual(digest, allocator_manifest["result_keys_sha256"])
        self.assertEqual(digest, allocator_policy["expected_result_keys_sha256"])

        changed = copy.deepcopy(keys)
        changed[0] = (changed[0][0], changed[0][1], "cold-requested")
        self.assertNotEqual(
            perf_compare.result_key_manifest_sha256(changed), digest
        )

    def test_allocator_filesystem_policy_rejects_child_and_source_identity_mutations(self):
        comparison_policy = allocator_policy_fixture()
        baseline = allocator_report()
        current = allocator_report(revision="current")
        perf_compare.compare_reports(baseline, current, comparison_policy)

        missing_child_process_id = copy.deepcopy(current)
        missing_child_process_id["filesystem_evidence"][0]["samples"][0].pop(
            "child_process_id"
        )
        with self.assertRaisesRegex(
            perf_compare.ComparisonInputError, "child_process_id"
        ):
            perf_compare.compare_reports(
                baseline, missing_child_process_id, comparison_policy
            )

        reused_child_process_id = copy.deepcopy(current)
        samples = reused_child_process_id["filesystem_evidence"][0]["samples"]
        samples[1]["child_process_id"] = samples[0]["child_process_id"]
        with self.assertRaisesRegex(
            perf_compare.ComparisonInputError, "must be unique"
        ):
            perf_compare.compare_reports(baseline, reused_child_process_id, comparison_policy)

    def test_allocator_child_identity_is_opt_in_for_legacy_policies(self):
        comparison_policy = allocator_policy_fixture()
        comparison_policy.pop("require_sample_child_identity")
        comparison_policy["expected_configuration"][
            "filesystem_fresh_child_per_sample"
        ] = False
        comparison_policy["expected_configuration"]["filesystem_process_isolated"] = (
            False
        )
        baseline = allocator_report()
        current = allocator_report(revision="current")
        for report_value in (baseline, current):
            report_value["configuration"]["filesystem_fresh_child_per_sample"] = False
            report_value["configuration"]["filesystem_process_isolated"] = False
            for evidence in report_value["filesystem_evidence"]:
                for sample in evidence["samples"]:
                    sample.pop("child_process_id")
        result = perf_compare.compare_reports(baseline, current, comparison_policy)
        self.assertEqual(result["status"], "pass")

        isolated_policy = allocator_policy_fixture()
        isolated_policy.pop("require_sample_child_identity")
        isolated_current = allocator_report(revision="current")
        isolated_current["filesystem_evidence"][0]["samples"][0].pop(
            "child_process_id"
        )
        with self.assertRaisesRegex(
            perf_compare.ComparisonInputError, "child_process_id"
        ):
            perf_compare.compare_reports(
                allocator_report(), isolated_current, isolated_policy
            )

        for field in (
            "filesystem_fresh_child_per_sample",
            "filesystem_process_isolated",
            "filesystem_root_selected",
        ):
            invalid_configuration = allocator_policy_fixture()
            invalid_configuration["expected_configuration"][field] = 1
            with self.subTest(expected_configuration_field=field):
                with self.assertRaisesRegex(
                    perf_compare.ComparisonInputError,
                    f"expected_configuration.{field} must be boolean",
                ):
                    perf_compare.validate_policy(invalid_configuration)

        invalid_report = allocator_report(revision="current")
        invalid_report["configuration"]["filesystem_process_isolated"] = 1
        with self.assertRaisesRegex(
            perf_compare.ComparisonInputError,
            "configuration.filesystem_process_isolated does not match policy",
        ):
            perf_compare.compare_reports(
                allocator_report(), invalid_report, allocator_policy_fixture()
            )

        invalid_policy = allocator_policy_fixture()
        invalid_policy["require_sample_child_identity"] = "yes"
        with self.assertRaisesRegex(
            perf_compare.ComparisonInputError,
            "require_sample_child_identity must be boolean",
        ):
            perf_compare.validate_policy(invalid_policy)

    def test_selected_root_xlsx_allocator_policy_rejects_identity_mutations(self):
        comparison_policy, corpus = xlsx_allocator_policy_fixture()
        baseline = xlsx_allocator_report(corpus)
        current = xlsx_allocator_report(corpus, revision="current")
        self.assertEqual(
            perf_compare.compare_reports(baseline, current, comparison_policy)["status"],
            "pass",
        )

        mutations = (
            (
                lambda report: report["filesystem_evidence"][0]["samples"][0].pop(
                    "child_process_id"
                ),
                "child_process_id",
            ),
            (
                lambda report: report["filesystem_evidence"][0]["samples"][1].update(
                    child_process_id=report["filesystem_evidence"][0]["samples"][0][
                        "child_process_id"
                    ]
                ),
                "must be unique",
            ),
            (
                lambda report: report["filesystem_evidence"][0]["corpus"].update(
                    archive_sha256="f" * 64
                ),
                "corpus identity",
            ),
            (
                lambda report: report["filesystem_evidence"][0]["samples"][0].update(
                    xlsx_source_sha256="f" * 64
                ),
                "xlsx_source_sha256",
            ),
            (
                lambda report: report["environment"].update(filesystem_type="ext4"),
                "filesystem identity mismatch",
            ),
            (
                lambda report: report["environment"].pop("storage_identifier"),
                "storage_identifier is required",
            ),
        )
        for mutate, message in mutations:
            with self.subTest(message=message):
                mutated = copy.deepcopy(current)
                mutate(mutated)
                with self.assertRaisesRegex(
                    perf_compare.ComparisonInputError, message
                ):
                    perf_compare.compare_reports(
                        baseline, mutated, comparison_policy
                    )

    def test_xlsx_allocator_policy_compares_retained_0271_abba_pairs(self):
        repository = Path(__file__).resolve().parents[1]
        evidence_root = (
            repository
            / "docs/performance/results/0271-xlsx-repeated-store-allocator-probe-20260824"
        )
        package_manifest = json.loads(
            (
                evidence_root
                / "0271-xlsx-repeated-store-allocator-probe-manifest.json"
            ).read_text(encoding="utf-8")
        )
        self.assertIs(type(package_manifest["schema_version"]), int)
        self.assertEqual(package_manifest["schema_version"], 1)
        self.assertIs(type(package_manifest["date"]), str)
        self.assertEqual(package_manifest["date"], "2026-08-24")
        self.assertEqual(
            package_manifest["change_id"],
            "0271-xlsx-repeated-store-allocator-probe",
        )
        self.assertEqual(package_manifest["status"], "exploratory_no_claim")
        expected_artifacts = {
            *(f"{role}.json" for role in ("a1", "a2", "b1", "b2")),
            "a1-b1-comparison.json",
            "a2-b2-comparison.json",
            "a1-b1-summary.txt",
            "a2-b2-summary.txt",
        }
        expected_artifact_roles = {
            "a1.json": "raw_control",
            "a2.json": "raw_control",
            "b1.json": "raw_candidate",
            "b2.json": "raw_candidate",
            "a1-b1-comparison.json": "comparison",
            "a2-b2-comparison.json": "comparison",
            "a1-b1-summary.txt": "summary",
            "a2-b2-summary.txt": "summary",
        }
        tracked_artifacts = package_manifest["tracked_artifacts"]
        self.assertIs(type(tracked_artifacts), list)
        tracked_paths = [artifact["path"] for artifact in tracked_artifacts]
        self.assertEqual(len(tracked_paths), len(set(tracked_paths)))
        self.assertEqual(len(tracked_paths), len(expected_artifacts))
        self.assertEqual(set(tracked_paths), expected_artifacts)
        artifacts = {
            artifact["path"]: artifact
            for artifact in tracked_artifacts
        }
        self.assertEqual(set(artifacts), expected_artifacts)
        self.assertEqual(
            {path: artifact["role"] for path, artifact in artifacts.items()},
            expected_artifact_roles,
        )
        for relative_path, artifact in artifacts.items():
            payload = (evidence_root / relative_path).read_bytes()
            self.assertEqual(len(payload), artifact["bytes"], relative_path)
            self.assertEqual(
                hashlib.sha256(payload).hexdigest(), artifact["sha256"], relative_path
            )

        protocol = package_manifest["protocol"]
        self.assertEqual(protocol["order"], ["a1", "b1", "b2", "a2"])
        self.assertEqual(protocol["samples_per_case"], 30)
        self.assertEqual(protocol["warmup_iterations_per_case"], 3)
        self.assertEqual(protocol["filesystem_cache_states"], ["warm"])
        for field in (
            "fresh_child_per_sample",
            "process_isolated",
            "filesystem_root_selected",
        ):
            self.assertIs(type(protocol[field]), bool)
            self.assertIs(protocol[field], True)
        self.assertEqual(protocol["filesystem_type"], "tmpfs")
        self.assertEqual(protocol["cpu_affinity"], "2")
        self.assertEqual(protocol["execution_workers"], [1])
        self.assertEqual(
            protocol["allocator"], "CountingSystemAllocator(std::alloc::System)"
        )
        self.assertEqual(
            protocol["instrumentation"], "system_allocator_operation_scoped"
        )
        self.assertEqual(
            protocol["allocation_scope"], "operation_global_system_allocator"
        )
        self.assertEqual(
            package_manifest["environment"]["source_destination_path_identity"],
            "unavailable",
        )
        self.assertEqual(
            package_manifest["environment"]["device_identity"], "unavailable"
        )
        self.assertIs(
            type(package_manifest["environment"]["source_destination_same_device"]),
            bool,
        )
        self.assertIs(
            package_manifest["environment"]["source_destination_same_device"], True
        )
        self.assertIsNone(package_manifest["environment"]["storage_identifier"])
        policy = json.loads(
            (
                repository
                / "docs/performance/perf-regression-policy-xlsx-allocator-v1.json"
            ).read_text(encoding="utf-8")
        )
        manifest_environment = package_manifest["environment"]
        self.assertEqual(
            set(manifest_environment),
            {
                "rustc_version",
                "target_os",
                "target_arch",
                "logical_cpus_available",
                "source_destination_same_device",
                "source_destination_path_identity",
                "device_identity",
                "storage_identifier",
            },
        )
        self.assertEqual(
            package_manifest["revisions"]["control"]["legs"], ["a1", "a2"]
        )
        self.assertEqual(
            package_manifest["revisions"]["candidate"]["legs"], ["b1", "b2"]
        )
        self.assertEqual(
            set(package_manifest["binaries"]["control"]["mode_bits_by_leg"]),
            {"a1", "a2"},
        )
        self.assertEqual(
            set(package_manifest["binaries"]["candidate"]["mode_bits_by_leg"]),
            {"b1", "b2"},
        )
        manifest_corpora = {
            corpus["case"]: corpus for corpus in package_manifest["corpora"]
        }
        self.assertEqual(
            [corpus["case"] for corpus in package_manifest["corpora"]],
            policy["required_cases"],
        )
        self.assertEqual(set(manifest_corpora), set(policy["required_cases"]))
        self.assertIs(
            type(package_manifest["comparison_policy"]["paired_results_identical"]),
            bool,
        )
        self.assertIs(
            package_manifest["comparison_policy"]["paired_results_identical"], True
        )
        self.assertEqual(
            package_manifest["comparison_policy"],
            {
                "metric_scope": "operation_metrics.allocation",
                "maximum_regression_percent": 5.0,
                "compared_metric_count_per_pair": 20,
                "regressions_per_pair": 0,
                "paired_results_identical": True,
                "latency": "excluded",
                "rss": "descriptive_only",
            },
        )
        self.assertEqual(
            package_manifest["claim_scope"],
            {
                "allocation": "exploratory_operation_scoped_observation_only",
                "excluded": [
                    "latency",
                    "operation_local_peak_memory",
                    "operation_local_peak_rss",
                    "physical_io",
                    "decompression",
                    "copy",
                    "broad_xlsx_performance",
                    "real_producer_breadth",
                ],
                "default_case_count": 36,
                "default_record_count": 198,
                "claim_0269": "retained_latency_only",
                "claim_registry_updated": False,
                "historical_classification_tables_updated": False,
            },
        )
        reports = {
            role: json.loads(
                (evidence_root / f"{role}.json").read_text(encoding="utf-8")
            )
            for role in ("a1", "a2", "b1", "b2")
        }
        perf_compare.validate_policy(policy)
        self.assertEqual(policy["minimum_samples"], 30)
        self.assertEqual(policy["expected_configuration"]["samples_per_case"], 30)
        self.assertEqual(
            policy["filesystem_identity_fields"],
            [
                "filesystem_type",
                "source_destination_same_device",
                "storage_identifier",
            ],
        )
        allocation_policy_class = next(
            item
            for item in policy["metric_classes"]
            if item["name"] == "allocation"
        )
        self.assertEqual(
            allocation_policy_class["max_regression_percent"],
            package_manifest["comparison_policy"]["maximum_regression_percent"],
        )
        self.assertEqual(allocation_policy_class["presence"], "required")
        self.assertTrue(
            all(
                pattern.startswith("operation_metrics/allocation/")
                for pattern in allocation_policy_class["path_globs"]
            )
        )
        for role, report in reports.items():
            revision_group = "control" if role.startswith("a") else "candidate"
            self.assertEqual(
                report["environment"]["git_revision"],
                package_manifest["revisions"][revision_group]["git_revision"],
            )
            expected_binary = package_manifest["binaries"][revision_group]
            binary = report["binary_identity"]
            # The manifest records one representative path per binary. A2 uses
            # a different absolute staging path for the same pinned binary, so
            # path equality is not claimed without a per-leg manifest field.
            self.assertTrue(Path(binary["path"]).is_absolute())
            self.assertEqual(binary["binary_sha256"], expected_binary["sha256"])
            self.assertEqual(binary["binary_bytes"], expected_binary["bytes"])
            self.assertEqual(
                binary["mode_bits"], expected_binary["mode_bits_by_leg"][role]
            )
            self.assertIs(type(binary["executable"]), bool)
            self.assertIs(binary["executable"], True)
            self.assertEqual(report["tool"]["target_os"], manifest_environment["target_os"])
            self.assertEqual(
                report["tool"]["target_arch"], manifest_environment["target_arch"]
            )
            self.assertEqual(
                report["tool"]["instrumentation"], protocol["instrumentation"]
            )
            environment = report["environment"]
            self.assertEqual(environment["rustc_version"], manifest_environment["rustc_version"])
            self.assertEqual(
                environment["logical_cpus_available"],
                manifest_environment["logical_cpus_available"],
            )
            self.assertIs(type(environment["source_destination_same_device"]), bool)
            self.assertIs(
                environment["source_destination_same_device"],
                manifest_environment["source_destination_same_device"],
            )
            self.assertEqual(
                environment["storage_identifier"],
                manifest_environment["storage_identifier"],
            )
            self.assertEqual(environment["filesystem_type"], protocol["filesystem_type"])
            self.assertEqual(environment["cpu_affinity"], protocol["cpu_affinity"])
            self.assertEqual(environment["allocator"], protocol["allocator"])
            # The retained schema has no root path/device identity; storage_identifier
            # is explicitly unavailable rather than synthesized by this policy.
            self.assertIsNone(environment["storage_identifier"])
            self.assertNotIn("filesystem_root_path", environment)
            self.assertNotIn("device_identifier", environment)
            configuration = report["configuration"]
            self.assertEqual(configuration["samples_per_case"], protocol["samples_per_case"])
            self.assertEqual(
                configuration["warmup_iterations_per_case"],
                protocol["warmup_iterations_per_case"],
            )
            self.assertEqual(
                configuration["filesystem_cache_states"],
                protocol["filesystem_cache_states"],
            )
            for field, protocol_field in (
                ("filesystem_fresh_child_per_sample", "fresh_child_per_sample"),
                ("filesystem_process_isolated", "process_isolated"),
                ("filesystem_root_selected", "filesystem_root_selected"),
            ):
                self.assertIs(type(configuration[field]), bool)
                self.assertIs(configuration[field], protocol[protocol_field])
            self.assertEqual(
                configuration["execution_workers"], protocol["execution_workers"]
            )
            for result_row in report["results"]:
                expected_corpus = manifest_corpora[result_row["case"]]
                for field, expected in expected_corpus.items():
                    if field in {"case", "selected_member", "selected_member_uncompressed_bytes"}:
                        continue
                    self.assertEqual(
                        result_row["corpus"][field], expected, f"{role}.{field}"
                    )
                self.assertEqual(
                    result_row["operation_metrics"]["latency_claim"],
                    "evidence_only_filesystem_selector",
                )
                self.assertEqual(
                    result_row["operation_metrics"]["allocation"]["scope"],
                    protocol["allocation_scope"],
                )
            evidence_by_case = {
                item["case"]: item for item in report["filesystem_evidence"]
            }
            self.assertEqual(set(evidence_by_case), set(manifest_corpora))
            for case, expected_corpus in manifest_corpora.items():
                evidence = evidence_by_case[case]
                result_corpus = next(
                    result_row["corpus"]
                    for result_row in report["results"]
                    if result_row["case"] == case
                )
                self.assertEqual(evidence["corpus"], result_corpus)
                for sample in evidence["samples"]:
                    self.assertEqual(
                        sample["xlsx_source_sha256"],
                        expected_corpus["archive_sha256"],
                    )
                    repeat_store = sample["xlsx_repeat_store"]
                    self.assertEqual(
                        repeat_store["selected_member"],
                        expected_corpus["selected_member"],
                    )
                    self.assertEqual(
                        repeat_store["selected_member_uncompressed_bytes"],
                        expected_corpus["selected_member_uncompressed_bytes"],
                    )
        canonical_corpora = {
            role: {
                row["case"]: json.dumps(
                    row["corpus"], sort_keys=True, separators=(",", ":")
                )
                for row in report["results"]
            }
            for role, report in reports.items()
        }
        for role in ("a2", "b1", "b2"):
            self.assertEqual(canonical_corpora[role], canonical_corpora["a1"])
        claim_scope = package_manifest["claim_scope"]
        self.assertEqual(
            claim_scope["allocation"],
            "exploratory_operation_scoped_observation_only",
        )
        self.assertIn("latency", claim_scope["excluded"])
        for field in (
            "claim_registry_updated",
            "historical_classification_tables_updated",
        ):
            self.assertIs(type(claim_scope[field]), bool)
            self.assertIs(claim_scope[field], False)
        self.assertNotEqual(
            reports["a1"]["binary_identity"]["mode_bits"],
            reports["a2"]["binary_identity"]["mode_bits"],
        )
        for baseline_role, current_role in (("a1", "b1"), ("a2", "b2")):
            with self.subTest(pair=f"{baseline_role}/{current_role}"):
                result = perf_compare.compare_reports(
                    reports[baseline_role], reports[current_role], policy
                )
                tracked_comparison = json.loads(
                    (
                        evidence_root
                        / f"{baseline_role}-{current_role}-comparison.json"
                    ).read_text(encoding="utf-8")
                )
                self.assertEqual(tracked_comparison["tool"]["version"], "1.3.2")
                historical_result = copy.deepcopy(result)
                historical_result["tool"]["version"] = "1.3.2"
                self.assertEqual(tracked_comparison, historical_result)
                self.assertEqual(
                    len(tracked_comparison["comparisons"]),
                    package_manifest["comparison_policy"][
                        "compared_metric_count_per_pair"
                    ],
                )
                self.assertEqual(
                    {
                        item["metric_class"]
                        for item in tracked_comparison["comparisons"]
                    },
                    {"allocation"},
                )
                self.assertEqual(
                    {
                        item["max_regression_percent"]
                        for item in tracked_comparison["comparisons"]
                    },
                    {
                        package_manifest["comparison_policy"][
                            "maximum_regression_percent"
                        ]
                    },
                )
                self.assertTrue(
                    all(
                        item["metric"].startswith("operation_metrics.allocation.")
                        for item in tracked_comparison["comparisons"]
                    )
                )
                self.assertEqual(
                    tracked_comparison["policy"],
                    {
                        "schema_version": policy["schema_version"],
                        "policy_id": policy["policy_id"],
                        "minimum_samples": policy["minimum_samples"],
                    },
                )
                tracked_summary = (
                    evidence_root
                    / f"{baseline_role}-{current_role}-summary.txt"
                ).read_text(encoding="utf-8")
                self.assertEqual(
                    tracked_summary.strip(),
                    "\n".join(
                        (
                            "PASS: 2 matched results, 20 metrics, 0 regressions",
                            "Latency comparison excluded for 2 evidence-only result(s)",
                        )
                    ),
                )
                self.assertEqual(result["status"], "pass")
                self.assertEqual(
                    result["summary"]["compared_metrics"],
                    package_manifest["comparison_policy"][
                        "compared_metric_count_per_pair"
                    ],
                )
                self.assertEqual(
                    result["summary"]["regressions"],
                    package_manifest["comparison_policy"]["regressions_per_pair"],
                )
                self.assertEqual(
                    result["summary"]["latency_claims"], "withheld_instrumentation"
                )
                self.assertEqual(result["summary"]["latency_compared_results"], 0)
                self.assertEqual(result["summary"]["matched_results"], 2)
                self.assertEqual(result["summary"]["latency_compared_results"], 0)
                self.assertEqual(result["summary"]["latency_excluded_results"], 2)

    def test_pass_compares_latency_and_available_resource_counters(self):
        baseline = report()
        current = report(104, "current")
        current["results"][0]["metrics"]["allocation_calls"] = [1020] * 5
        current["results"][0]["metrics"]["peak_rss_bytes"] = [10_300] * 5
        current["results"][0]["metrics"]["instructions"] = [52_000] * 5
        result = perf_compare.compare_reports(baseline, current, policy())
        self.assertEqual(result["status"], "pass")
        self.assertEqual(result["summary"]["matched_results"], 1)
        self.assertEqual(result["summary"]["compared_metrics"], 7)
        self.assertFalse(result["regressions"])

    def test_descriptive_parallel_envelope_is_shape_checked_but_not_compared(self):
        result = perf_compare.compare_reports(
            descriptive_parallel_report(),
            descriptive_parallel_report(revision="current"),
            policy(),
        )
        self.assertEqual(result["status"], "pass")
        self.assertEqual(result["summary"]["compared_metrics"], 7)

    def test_parallel_lock_diagnostics_expose_direct_mutex_p50(self):
        baseline = descriptive_parallel_lock_report()
        current = descriptive_parallel_lock_report(revision="current")
        perf_compare.validate_parallel_metrics(baseline)
        result = perf_compare.compare_reports(baseline, current, policy())
        self.assertEqual(result["status"], "pass")
        self.assertEqual(
            baseline["parallel_metrics"]["cases"][0]["lock_wait_ns"]["value"],
            7,
        )

    def test_parallel_lock_diagnostics_cross_checks_scalar_and_totals(self):
        current = descriptive_parallel_lock_report(revision="current")
        current["parallel_metrics"]["cases"][0]["lock_wait_ns"]["value"] = 6
        with self.assertRaisesRegex(
            perf_compare.ComparisonInputError,
            "lock_wait_ns.value does not match",
        ):
            perf_compare.validate_parallel_metrics(current)

        current = descriptive_parallel_lock_report(revision="current")
        current["results"][0]["source"]["opc_cache"]["lock_diagnostics"][
            "total_lock_wait_ns"
        ][0] = 10
        with self.assertRaisesRegex(
            perf_compare.ComparisonInputError,
            "total_lock_wait_ns does not equal cache plus flight",
        ):
            perf_compare.validate_parallel_metrics(current)

        current = descriptive_parallel_lock_report(revision="current")
        current["results"][0]["case"] = "opc_source_cache_control_contention"
        current["parallel_metrics"]["cases"][0]["case"] = (
            "opc_source_cache_control_contention"
        )
        diagnostics = current["results"][0]["source"]["opc_cache"][
            "lock_diagnostics"
        ]
        diagnostics["cache_lock_acquisitions"][0] = 0
        diagnostics["flight_lock_acquisitions"][0] = 0
        diagnostics["total_lock_acquisitions"][0] = 0
        with self.assertRaisesRegex(
            perf_compare.ComparisonInputError,
            "total_lock_acquisitions must be positive",
        ):
            perf_compare.validate_parallel_metrics(current)

    def test_parallel_lock_diagnostics_are_required_for_instrumented_contention_rows(self):
        current = descriptive_parallel_lock_report(revision="current")
        current["results"][0]["case"] = "opc_source_cache_control_contention"
        current["parallel_metrics"]["cases"][0]["case"] = (
            "opc_source_cache_control_contention"
        )
        del current["results"][0]["source"]
        with self.assertRaisesRegex(
            perf_compare.ComparisonInputError,
            r"source\.opc_cache is required",
        ):
            perf_compare.validate_parallel_metrics(current)

        current = descriptive_parallel_lock_report(revision="current")
        current["results"][0]["case"] = "opc_source_cache_control_contention"
        current["parallel_metrics"]["cases"][0]["case"] = (
            "opc_source_cache_control_contention"
        )
        del current["results"][0]["source"]["opc_cache"]["lock_diagnostics"]
        with self.assertRaisesRegex(
            perf_compare.ComparisonInputError,
            "lock_diagnostics is required",
        ):
            perf_compare.validate_parallel_metrics(current)

    def test_parallel_lock_diagnostics_reject_boolean_worker_team_count(self):
        current = descriptive_parallel_lock_report(revision="current")
        current["results"][0]["source"]["opc_cache"][
            "persistent_worker_teams_created"
        ] = True
        with self.assertRaisesRegex(
            perf_compare.ComparisonInputError,
            "persistent_worker_teams_created must be a non-negative",
        ):
            perf_compare.validate_parallel_metrics(current)

    def test_parallel_envelope_coexists_with_strict_operation_metrics(self):
        baseline = descriptive_parallel_report()
        current = descriptive_parallel_report(revision="current")
        for report_value in (baseline, current):
            report_value["results"][0]["operation_metrics"] = (
                operation_metrics_report_fields()
            )
        result = perf_compare.compare_reports(
            baseline, current, self.operation_metrics_policy()
        )
        self.assertEqual(result["status"], "pass")
        self.assertEqual(result["summary"]["compared_metrics"], 14)

    def test_parallel_envelope_metadata_and_sample_order_fail_closed(self):
        current = descriptive_parallel_report(revision="current")
        current["parallel_metrics"]["claim"] = "latency"
        with self.assertRaisesRegex(
            perf_compare.ComparisonInputError, "parallel_metrics.claim"
        ):
            perf_compare.compare_reports(
                descriptive_parallel_report(), current, policy()
            )

        current = descriptive_parallel_report(revision="current")
        current["results"][0]["execution"]["worker_count"] = 8
        with self.assertRaisesRegex(
            perf_compare.ComparisonInputError, "worker_count=8.*execution_workers"
        ):
            perf_compare.compare_reports(
                descriptive_parallel_report(), current, policy()
            )

        current = descriptive_parallel_report(revision="current")
        current["parallel_metrics"]["configured_worker_budget"]["value"] = [1, 4]
        with self.assertRaisesRegex(
            perf_compare.ComparisonInputError,
            "configured_worker_budget.value must match",
        ):
            perf_compare.compare_reports(
                descriptive_parallel_report(), current, policy()
            )

        current = descriptive_parallel_report(revision="current")
        current["parallel_metrics"]["cases"][0]["configured_worker_count"][
            "value"
        ] = 1
        with self.assertRaisesRegex(
            perf_compare.ComparisonInputError,
            "configured_worker_count.value does not match",
        ):
            perf_compare.compare_reports(
                descriptive_parallel_report(), current, policy()
            )

        current = descriptive_parallel_report(revision="current")
        current["parallel_metrics"]["cases"][0]["configured_worker_count"][
            "value"
        ] = [2]
        with self.assertRaisesRegex(
            perf_compare.ComparisonInputError,
            "must be a non-negative unsigned 64-bit integer scalar",
        ):
            perf_compare.compare_reports(
                descriptive_parallel_report(), current, policy()
            )

        current = descriptive_parallel_report(revision="current")
        current["results"][0]["execution"]["logical_tasks"] = 6
        with self.assertRaisesRegex(
            perf_compare.ComparisonInputError, "deterministic_task_count.value does not match"
        ):
            perf_compare.compare_reports(
                descriptive_parallel_report(), current, policy()
            )

        current = descriptive_parallel_report(revision="current")
        current["results"][0]["source"]["simulation"]["physical_request_count"][
            0
        ] = 99
        with self.assertRaisesRegex(
            perf_compare.ComparisonInputError,
            "deterministic_chunk_count.value does not match",
        ):
            perf_compare.compare_reports(
                descriptive_parallel_report(), current, policy()
            )

        current = descriptive_parallel_report(revision="current")
        current["parallel_metrics"]["cases"][0]["cache_state"] = "warm"
        with self.assertRaisesRegex(
            perf_compare.ComparisonInputError, "cache_state is present"
        ):
            perf_compare.compare_reports(
                descriptive_parallel_report(), current, policy()
            )

        current = descriptive_parallel_report(revision="current")
        current["results"][0]["source"]["opc_cache"] = {
            "worker_count": 2,
            "persistent_worker_teams_created": 1,
        }
        with self.assertRaisesRegex(
            perf_compare.ComparisonInputError,
            "observed_local_worker_count.status",
        ):
            perf_compare.compare_reports(
                descriptive_parallel_report(), current, policy()
            )

        current = descriptive_parallel_report(revision="current")
        current["results"][0]["elapsed_ns"]["samples"] = [200, 100, 100, 100, 100]
        with self.assertRaisesRegex(
            perf_compare.ComparisonInputError, "samples must be sorted"
        ):
            perf_compare.compare_reports(
                descriptive_parallel_report(), current, policy()
            )

        current = descriptive_parallel_report(revision="current")
        current["results"][0]["elapsed_ns"]["samples"] = [100, 100, 102, 103, 104]
        with self.assertRaisesRegex(
            perf_compare.ComparisonInputError, "sample_order.*sorted sample identity"
        ):
            perf_compare.compare_reports(
                descriptive_parallel_report(), current, policy()
            )

        current = descriptive_parallel_report(revision="current")
        del current["parallel_metrics"]
        with self.assertRaisesRegex(
            perf_compare.ComparisonInputError, "both emit parallel_metrics"
        ):
            perf_compare.compare_reports(
                descriptive_parallel_report(), current, policy()
            )

        current = descriptive_parallel_report(revision="current")
        current["results"][0]["elapsed_ns"]["sample_order"] = [0, 0, 1, 2, 3]
        with self.assertRaisesRegex(
            perf_compare.ComparisonInputError, "sample_order.*permutation"
        ):
            perf_compare.compare_reports(
                descriptive_parallel_report(), current, policy()
            )

        current = descriptive_parallel_report(revision="current")
        current["parallel_metrics"]["cases"][0]["deterministic_chunk_count"][
            "value"
        ] = [1]
        with self.assertRaisesRegex(
            perf_compare.ComparisonInputError, "deterministic_chunk_count.value"
        ):
            perf_compare.compare_reports(
                descriptive_parallel_report(), current, policy()
            )

    def test_additive_sink_histogram_field_is_comparator_compatible(self):
        current = report(revision="current")
        current["results"][0]["sink"] = {
            "accepted_bytes": 3,
            "write_calls": 2,
            "largest_write": 2,
            "write_size_buckets": {
                "bytes_0": 0,
                "bytes_1_to_512": 2,
                "bytes_513_to_4096": 0,
                "bytes_4097_to_16384": 0,
                "bytes_16385_to_65536": 0,
                "bytes_over_65536": 0,
            },
        }
        sink = current["results"][0]["sink"]
        self.assertEqual(sum(sink["write_size_buckets"].values()), sink["write_calls"])
        self.assertEqual(sink["accepted_bytes"], 3)
        self.assertEqual(sink["largest_write"], 2)
        result = perf_compare.compare_reports(report(), current, policy())
        self.assertEqual(result["status"], "pass")
        self.assertEqual(result["summary"]["compared_metrics"], 7)

    def test_operation_metric_vectors_unwrap_and_regressions_fail(self):
        baseline = report()
        current = report(revision="current")
        baseline["results"][0]["operation_metrics"] = (
            operation_metrics_report_fields()
        )
        current["results"][0]["operation_metrics"] = (
            operation_metrics_report_fields()
        )

        result = perf_compare.compare_reports(
            baseline, current, self.operation_metrics_policy()
        )
        self.assertEqual(result["status"], "pass")
        self.assertEqual(result["summary"]["compared_metrics"], 14)

        current["results"][0]["operation_metrics"]["sink"]["write_calls"][
            "values"
        ] = [3] * 5
        current["results"][0]["operation_metrics"]["process"]["read_bytes"][
            "values"
        ] = [5] * 5
        current["results"][0]["operation_metrics"]["source"][
            "logical_read_requested_bytes"
        ]["values"] = [30] * 5
        current["results"][0]["operation_metrics"]["source"][
            "logical_read_returned_bytes"
        ]["values"] = [30] * 5
        current["results"][0]["operation_metrics"]["source"][
            "max_concurrent_reads"
        ]["values"] = [2] * 5
        result = perf_compare.compare_reports(
            baseline, current, self.operation_metrics_policy()
        )
        self.assertEqual(result["status"], "regression")
        self.assertTrue(
            {
                "operation_metrics.process.read_bytes",
                "operation_metrics.sink.write_calls",
                "operation_metrics.source.logical_read_requested_bytes",
                "operation_metrics.source.logical_read_returned_bytes",
                "operation_metrics.source.max_concurrent_reads",
            }
            <= {item["metric"] for item in result["regressions"]}
        )

    def test_filesystem_evidence_latency_is_excluded_from_comparison(self):
        baseline = report(value=100)
        current = report(value=200, revision="current")
        for item in (baseline["results"][0], current["results"][0]):
            item["operation_metrics"] = operation_metrics_report_fields()
            item["operation_metrics"][
                "latency_claim"
            ] = perf_compare.EVIDENCE_ONLY_LATENCY_CLAIM
        result = perf_compare.compare_reports(
            baseline, current, self.operation_metrics_policy()
        )
        self.assertEqual(result["status"], "pass")
        self.assertEqual(result["summary"]["latency_compared_results"], 0)
        self.assertEqual(result["summary"]["latency_excluded_results"], 1)
        self.assertNotIn(
            "elapsed_ns.p50", {item["metric"] for item in result["comparisons"]}
        )

    def test_latency_claim_mismatch_fails_closed(self):
        baseline = report()
        current = report(revision="current")
        baseline["results"][0]["operation_metrics"] = operation_metrics_report_fields()
        current["results"][0]["operation_metrics"] = operation_metrics_report_fields()
        for item in (baseline, current):
            source = item["results"][0]["operation_metrics"]["source"]
            source["status"] = "not_applicable"
            source["counter_scope"] = "not_applicable_in_process_sink"
            for field, vector in source.items():
                if field in {"status", "counter_scope"}:
                    continue
                vector["status"] = "not_applicable"
                vector.pop("values", None)
        current["results"][0]["operation_metrics"][
            "latency_claim"
        ] = perf_compare.COMPARABLE_LATENCY_CLAIM
        with self.assertRaisesRegex(
            perf_compare.ComparisonInputError, "latency claim mismatch"
        ):
            perf_compare.compare_reports(
                baseline, current, self.operation_metrics_policy()
            )

    def test_source_counter_scope_mismatch_fails_closed(self):
        baseline = report()
        current = report(revision="current")
        baseline["results"][0]["operation_metrics"] = operation_metrics_report_fields()
        current["results"][0]["operation_metrics"] = operation_metrics_report_fields()
        source = current["results"][0]["operation_metrics"]["source"]
        source["status"] = "not_applicable"
        source["counter_scope"] = "untimed_source_replay_only"
        for field, vector in source.items():
            if field in {"status", "counter_scope"}:
                continue
            vector["status"] = "not_applicable"
            vector.pop("values", None)
        with self.assertRaisesRegex(
            perf_compare.ComparisonInputError, "source counter scope mismatch"
        ):
            perf_compare.compare_reports(
                baseline, current, self.operation_metrics_policy()
            )

    def test_filesystem_xlsx_counter_scope_is_valid_for_not_applicable_source(self):
        baseline = report()
        current = report(revision="current")
        for item in (baseline["results"][0], current["results"][0]):
            item["operation_metrics"] = (
                filesystem_xlsx_operation_metrics_report_fields()
            )
            item["elapsed_ns"]["sample_order"] = list(range(5))

        result = perf_compare.compare_reports(
            baseline, current, self.operation_metrics_policy()
        )
        self.assertEqual(result["status"], "pass")

    def test_filesystem_xlsx_counter_scope_requires_sample_order(self):
        for missing_value in ("missing", None):
            with self.subTest(sample_order=missing_value):
                baseline = report()
                current = report(revision="current")
                for item in (baseline["results"][0], current["results"][0]):
                    item["operation_metrics"] = (
                        filesystem_xlsx_operation_metrics_report_fields()
                    )
                    if missing_value != "missing":
                        item["elapsed_ns"]["sample_order"] = missing_value
                with self.assertRaisesRegex(
                    perf_compare.ComparisonInputError,
                    "elapsed_ns.sample_order is required",
                ):
                    perf_compare.compare_reports(
                        baseline, current, self.operation_metrics_policy()
                    )

    def test_accounted_materialization_scope_is_evidence_only_and_fail_closed(self):
        baseline, current, comparison_policy = (
            opc_source_materialize_report_pair(
                accounted=True,
                comparison_policy=self.operation_metrics_policy(),
            )
        )
        result = perf_compare.compare_reports(baseline, current, comparison_policy)
        self.assertEqual(result["status"], "pass")
        self.assertEqual(result["summary"]["latency_compared_results"], 0)
        self.assertEqual(result["summary"]["latency_excluded_results"], 1)
        mutations = (
            lambda m: m.pop("opc_zip"),
            lambda m: m["opc_zip"].update(scope=perf_compare.OPC_ZIP_SCOPE),
            lambda m: m["opc_zip"]["output_bytes_accepted"].update(scope=perf_compare.OPC_ZIP_SCOPE),
            lambda m: m.update(latency_claim=perf_compare.OPC_ZIP_EVIDENCE_ONLY_LATENCY_CLAIM),
            lambda m: m.update(latency_claim=perf_compare.OPC_SOURCE_MATERIALIZATION_EVIDENCE_ONLY_LATENCY_CLAIM),
        )
        for index, mutate in enumerate(mutations):
            with self.subTest(index=index):
                malformed = copy.deepcopy(current)
                mutate(malformed["results"][0]["operation_metrics"])
                with self.assertRaises(perf_compare.ComparisonInputError):
                    perf_compare.compare_reports(baseline, malformed, comparison_policy)

    def test_opc_materialization_claims_are_bound_to_result_case(self):
        baseline, current, comparison_policy = opc_source_materialize_report_pair(
            comparison_policy=self.operation_metrics_policy()
        )
        current["results"][0]["elapsed_ns"]["sample_order"] = list(range(5))
        current["results"][0]["operation_metrics"] = (
            opc_source_materialize_accounted_operation_metrics_report_fields()
        )
        with self.assertRaisesRegex(
            perf_compare.ComparisonInputError,
            "latency_claim=.*requires result case .*opc_source_materialize_accounted",
        ):
            perf_compare.compare_reports(baseline, current, comparison_policy)

        baseline, current, comparison_policy = opc_source_materialize_report_pair(
            accounted=True,
            comparison_policy=self.operation_metrics_policy(),
        )
        current["results"][0]["operation_metrics"] = (
            opc_source_materialize_operation_metrics_report_fields()
        )
        with self.assertRaisesRegex(
            perf_compare.ComparisonInputError,
            "latency_claim=.*requires result case .*opc_source_materialize'",
        ):
            perf_compare.compare_reports(baseline, current, comparison_policy)

    def test_opc_materialization_oracle_identity_requires_new_accounted_marker(self):
        baseline, current, comparison_policy = opc_source_materialize_report_pair(
            accounted=True,
            comparison_policy=self.operation_metrics_policy(),
        )
        current["configuration"].pop(
            perf_compare.OPC_SOURCE_MATERIALIZE_ORACLE_CONFIG_FIELD
        )
        with self.assertRaisesRegex(
            perf_compare.ComparisonInputError,
            "configuration.opc_source_materialize_oracle is required",
        ):
            perf_compare.compare_reports(baseline, current, comparison_policy)

        baseline, current, comparison_policy = opc_source_materialize_report_pair(
            accounted=True,
            comparison_policy=self.operation_metrics_policy(),
        )
        current["configuration"][
            perf_compare.OPC_SOURCE_MATERIALIZE_ORACLE_CONFIG_FIELD
        ] = "future-oracle"
        with self.assertRaisesRegex(
            perf_compare.ComparisonInputError,
            "configuration.opc_source_materialize_oracle must be",
        ):
            perf_compare.compare_reports(baseline, current, comparison_policy)

        baseline, current, comparison_policy = opc_source_materialize_report_pair(
            comparison_policy=self.operation_metrics_policy()
        )
        self.assertNotIn(
            perf_compare.OPC_SOURCE_MATERIALIZE_ORACLE_CONFIG_FIELD,
            baseline["configuration"],
        )
        self.assertNotIn(
            perf_compare.OPC_SOURCE_MATERIALIZE_ORACLE_CONFIG_FIELD,
            current["configuration"],
        )
        result = perf_compare.compare_reports(baseline, current, comparison_policy)
        self.assertEqual(result["status"], "pass")

        for item in (baseline, current):
            item["configuration"][
                perf_compare.OPC_SOURCE_MATERIALIZE_ORACLE_CONFIG_FIELD
            ] = perf_compare.OPC_SOURCE_MATERIALIZE_ORACLE_VERSION
        result = perf_compare.compare_reports(baseline, current, comparison_policy)
        self.assertEqual(result["status"], "pass")

        invalid_legacy = copy.deepcopy(current)
        invalid_legacy["configuration"][
            perf_compare.OPC_SOURCE_MATERIALIZE_ORACLE_CONFIG_FIELD
        ] = "future-oracle"
        with self.assertRaisesRegex(
            perf_compare.ComparisonInputError,
            "configuration.opc_source_materialize_oracle must be",
        ):
            perf_compare.compare_reports(baseline, invalid_legacy, comparison_policy)

    def test_xlsx_cell_values_range_accounting_configuration_identity(self):
        baseline, current, comparison_policy = managed_xlsx_reports()
        field = perf_compare.XLSX_CELL_VALUES_RANGE_ACCOUNTING_CONFIG_FIELD
        version = perf_compare.XLSX_CELL_VALUES_RANGE_ACCOUNTING_VERSION
        self.assertEqual(
            perf_compare.compare_reports(baseline, current, comparison_policy)["status"],
            "pass",
        )
        current["configuration"][field] = version
        with self.assertRaisesRegex(
            perf_compare.ComparisonInputError, "benchmark configuration mismatch"
        ):
            perf_compare.compare_reports(baseline, current, comparison_policy)
        baseline["configuration"][field] = version
        self.assertEqual(
            perf_compare.compare_reports(baseline, current, comparison_policy)["status"],
            "pass",
        )
        for invalid in (None, False, 1, "", "future-accounting"):
            with self.subTest(invalid=invalid):
                malformed = copy.deepcopy(current)
                malformed["configuration"][field] = invalid
                with self.assertRaisesRegex(
                    perf_compare.ComparisonInputError,
                    "configuration.xlsx_cell_values_range_accounting must be",
                ):
                    perf_compare.compare_reports(baseline, malformed, comparison_policy)

    def test_xlsx_cell_values_range_accounting_is_bound_to_source_selectors(self):
        field = perf_compare.XLSX_CELL_VALUES_RANGE_ACCOUNTING_CONFIG_FIELD
        version = perf_compare.XLSX_CELL_VALUES_RANGE_ACCOUNTING_VERSION
        for cases in ([], ["opc_open"], ["xlsx_eager_cell_values_one_edit_save"],
                      ["xlsx_source_backed_cell_clear_edit_save"], [None], None):
            with self.subTest(cases=cases):
                with self.assertRaisesRegex(
                    perf_compare.ComparisonInputError,
                    "requires a source-backed XLSX scalar cell-values edit/save case",
                ):
                    perf_compare._validate_xlsx_cell_values_range_accounting(
                        {field: version, "cases": cases}, "current"
                    )
        for case in perf_compare.XLSX_CELL_VALUES_RANGE_ACCOUNTING_CASES:
            perf_compare._validate_xlsx_cell_values_range_accounting(
                {field: version, "cases": ["opc_open", case]}, "current"
            )

    def test_opc_source_materialization_scope_and_claim_are_evidence_only(self):
        baseline, current, comparison_policy = opc_source_materialize_report_pair(
            comparison_policy=self.operation_metrics_policy()
        )

        result = perf_compare.compare_reports(
            baseline, current, comparison_policy
        )
        self.assertEqual(result["status"], "pass")
        self.assertEqual(result["summary"]["latency_compared_results"], 0)
        self.assertEqual(result["summary"]["latency_excluded_results"], 1)
        self.assertNotIn(
            "elapsed_ns.p50", {item["metric"] for item in result["comparisons"]}
        )

    def test_opc_source_materialization_scope_requires_dedicated_claim(self):
        baseline, current, comparison_policy = opc_source_materialize_report_pair(
            comparison_policy=self.operation_metrics_policy()
        )
        current["results"][0]["operation_metrics"][
            "latency_claim"
        ] = perf_compare.EVIDENCE_ONLY_LATENCY_CLAIM
        with self.assertRaisesRegex(
            perf_compare.ComparisonInputError,
            "source.counter_scope=.*requires latency_claim",
        ):
            perf_compare.compare_reports(
                baseline, current, comparison_policy
            )

        baseline, current, comparison_policy = opc_source_materialize_report_pair(
            comparison_policy=self.operation_metrics_policy()
        )
        current["results"][0]["operation_metrics"]["source"][
            "counter_scope"
        ] = "timed_read_at"
        with self.assertRaisesRegex(
            perf_compare.ComparisonInputError,
            "latency_claim=.*requires source.counter_scope",
        ):
            perf_compare.compare_reports(
                baseline, current, comparison_policy
            )

    def test_opc_source_materialization_scope_rejects_comparable_claim(self):
        baseline, current, comparison_policy = opc_source_materialize_report_pair(
            comparison_policy=self.operation_metrics_policy()
        )
        current["results"][0]["operation_metrics"][
            "latency_claim"
        ] = perf_compare.COMPARABLE_LATENCY_CLAIM
        with self.assertRaisesRegex(
            perf_compare.ComparisonInputError,
            "source.counter_scope=.*requires latency_claim",
        ):
            perf_compare.compare_reports(
                baseline, current, comparison_policy
            )

    def test_source_counter_scope_status_compatibility_fails_closed(self):
        baseline = report()
        current = report(revision="current")
        baseline["results"][0]["operation_metrics"] = operation_metrics_report_fields()
        current["results"][0]["operation_metrics"] = operation_metrics_report_fields()
        current["results"][0]["operation_metrics"]["source"][
            "counter_scope"
        ] = "not_applicable_filesystem_xlsx"
        with self.assertRaisesRegex(
            perf_compare.ComparisonInputError,
            "source.status=.*incompatible.*not_applicable_filesystem_xlsx",
        ):
            perf_compare.compare_reports(
                baseline, current, self.operation_metrics_policy()
            )

        baseline = report()
        current = report(revision="current")
        for item in (baseline["results"][0], current["results"][0]):
            item["operation_metrics"] = filesystem_xlsx_operation_metrics_report_fields()
            item["elapsed_ns"]["sample_order"] = list(range(5))
        current["results"][0]["operation_metrics"]["source"][
            "counter_scope"
        ] = "timed_read_at"
        with self.assertRaisesRegex(
            perf_compare.ComparisonInputError,
            "source.status=.*incompatible.*timed_read_at",
        ):
            perf_compare.compare_reports(
                baseline, current, self.operation_metrics_policy()
            )

    def test_invalid_source_counter_scope_fails_closed(self):
        baseline = report()
        current = report(revision="current")
        baseline["results"][0]["operation_metrics"] = operation_metrics_report_fields()
        current["results"][0]["operation_metrics"] = operation_metrics_report_fields()
        current["results"][0]["operation_metrics"]["source"][
            "counter_scope"
        ] = "future_scope"
        with self.assertRaisesRegex(
            perf_compare.ComparisonInputError,
            "operation_metrics.source.counter_scope must be one of",
        ):
            perf_compare.compare_reports(
                baseline, current, self.operation_metrics_policy()
            )

    def test_measured_source_requires_evidence_only_latency_claim(self):
        baseline = report()
        current = report(revision="current")
        baseline["results"][0]["operation_metrics"] = operation_metrics_report_fields()
        current["results"][0]["operation_metrics"] = operation_metrics_report_fields()
        current["results"][0]["operation_metrics"][
            "latency_claim"
        ] = perf_compare.COMPARABLE_LATENCY_CLAIM
        with self.assertRaisesRegex(
            perf_compare.ComparisonInputError,
            "measured source metrics require",
        ):
            perf_compare.compare_reports(
                baseline, current, self.operation_metrics_policy()
            )

    def test_malformed_metric_vector_wrapper_fails_closed(self):
        baseline = report()
        current = report(revision="current")
        for item in (baseline["results"][0], current["results"][0]):
            item["operation_metrics"] = operation_metrics_report_fields()
            del item["operation_metrics"]["sink"]["write_calls"]["values"]
        with self.assertRaisesRegex(
            perf_compare.ComparisonInputError, "values is required"
        ):
            perf_compare.compare_reports(
                baseline, current, self.operation_metrics_policy()
            )

    def test_metric_vector_metadata_mismatches_fail_closed(self):
        def reports():
            baseline = report()
            current = report(revision="current")
            baseline["results"][0]["operation_metrics"] = (
                operation_metrics_report_fields()
            )
            current["results"][0]["operation_metrics"] = (
                operation_metrics_report_fields()
            )
            return baseline, current

        baseline, current = reports()
        current["results"][0]["operation_metrics"]["process"]["read_bytes"][
            "scope"
        ] = "procfs_operation_delta"
        with self.assertRaisesRegex(
            perf_compare.ComparisonInputError,
            "scope mismatch.*operation_metrics.process.read_bytes",
        ):
            perf_compare.compare_reports(
                baseline, current, self.operation_metrics_policy()
            )

        baseline, current = reports()
        current["results"][0]["operation_metrics"]["sink"]["output_bytes"][
            "status"
        ] = "unavailable"
        current["results"][0]["operation_metrics"]["sink"]["status"] = (
            "unavailable"
        )
        with self.assertRaisesRegex(
            perf_compare.ComparisonInputError,
            "status mismatch.*operation_metrics.sink.output_bytes",
        ):
            perf_compare.compare_reports(
                baseline, current, self.operation_metrics_policy()
            )

        baseline, current = reports()
        current_process = current["results"][0]["operation_metrics"]["process"]
        current_process["status"] = "unavailable"
        for key, value in current_process.items():
            if key == "status":
                continue
            value["status"] = "unavailable"
            del value["values"]
        with self.assertRaisesRegex(
            perf_compare.ComparisonInputError,
            "status mismatch.*operation_metrics.process",
        ):
            perf_compare.compare_reports(
                baseline, current, self.operation_metrics_policy()
            )

        baseline, current = reports()
        del current["results"][0]["operation_metrics"]["sink"]["output_bytes"]
        with self.assertRaisesRegex(
            perf_compare.ComparisonInputError,
            "operation_metrics.sink keys mismatch.*output_bytes",
        ):
            perf_compare.compare_reports(
                baseline, current, self.operation_metrics_policy()
            )

        baseline, current = reports()
        del current["results"][0]["operation_metrics"]["process"]["rchar"]
        with self.assertRaisesRegex(
            perf_compare.ComparisonInputError,
            "operation_metrics.process keys mismatch.*rchar",
        ):
            perf_compare.compare_reports(
                baseline, current, self.operation_metrics_policy()
            )

        baseline = report()
        current = report(revision="current")
        baseline["results"][0]["operation_metrics"] = operation_metrics_report_fields()
        current["results"][0]["operation_metrics"] = operation_metrics_report_fields()
        del current["results"][0]["operation_metrics"]["source"][
            "logical_read_calls"
        ]
        with self.assertRaisesRegex(
            perf_compare.ComparisonInputError,
            "operation_metrics.source keys mismatch.*logical_read_calls",
        ):
            perf_compare.compare_reports(
                baseline, current, self.operation_metrics_policy()
            )

    def test_partial_metric_vector_wrappers_fail_but_ordinary_dicts_pass(self):
        partial_wrappers = (
            {"values": [1] * 5},
            {"scope": "test_scope"},
            {"status": "measured"},
            {"status": "measured", "values": [1] * 5},
        )
        for partial in partial_wrappers:
            with self.subTest(partial=partial):
                baseline = report()
                current = report(revision="current")
                baseline["results"][0]["operation_metrics"] = (
                    operation_metrics_report_fields()
                )
                current["results"][0]["operation_metrics"] = (
                    operation_metrics_report_fields()
                )
                current["results"][0]["operation_metrics"]["sink"][
                    "output_bytes"
                ] = partial
                with self.assertRaisesRegex(
                    perf_compare.ComparisonInputError,
                    "partial MetricVector wrapper",
                ):
                    perf_compare.compare_reports(
                        baseline, current, self.operation_metrics_policy()
                    )

        baseline = report()
        current = report(revision="current")
        for item in (baseline["results"][0], current["results"][0]):
            item["ordinary_metadata"] = {
                "values": [1] * 5,
                "scope": "ordinary_scope",
                "status": "measured",
                "message": "ordinary dictionary",
            }
        result = perf_compare.compare_reports(baseline, current, policy())
        self.assertEqual(result["status"], "pass")
        self.assertEqual(result["summary"]["compared_metrics"], 7)

    def test_known_metric_vector_leaves_reject_markerless_and_unknown_shapes(self):
        malformed_leaves = ({}, {"foo": 1}, {"status": "ok"})
        for malformed in malformed_leaves:
            with self.subTest(malformed=malformed):
                baseline = report()
                current = report(revision="current")
                baseline["results"][0]["operation_metrics"] = (
                    operation_metrics_report_fields()
                )
                current["results"][0]["operation_metrics"] = (
                    operation_metrics_report_fields()
                )
                current["results"][0]["operation_metrics"]["sink"][
                    "output_bytes"
                ] = malformed
                with self.assertRaisesRegex(
                    perf_compare.ComparisonInputError,
                    "operation_metrics.sink.output_bytes",
                ):
                    perf_compare.compare_reports(
                        baseline, current, self.operation_metrics_policy()
                    )

        baseline = report()
        current = report(revision="current")
        baseline["results"][0]["operation_metrics"] = operation_metrics_report_fields()
        current["results"][0]["operation_metrics"] = operation_metrics_report_fields()
        current["results"][0]["operation_metrics"]["sink"]["mystery"] = 1
        with self.assertRaisesRegex(
            perf_compare.ComparisonInputError,
            "operation_metrics.sink keys mismatch.*mystery",
        ):
            perf_compare.compare_reports(
                baseline, current, self.operation_metrics_policy()
            )

        baseline = report()
        current = report(revision="current")
        baseline["results"][0]["operation_metrics"] = operation_metrics_report_fields()
        current["results"][0]["operation_metrics"] = operation_metrics_report_fields()
        current["results"][0]["operation_metrics"]["sink"]["write_calls"][
            "mystery"
        ] = 1
        with self.assertRaisesRegex(
            perf_compare.ComparisonInputError,
            "operation_metrics.sink.write_calls.*unknown keys.*mystery",
        ):
            perf_compare.compare_reports(
                baseline, current, self.operation_metrics_policy()
            )

        baseline = report()
        current = report(revision="current")
        baseline["results"][0]["operation_metrics"] = operation_metrics_report_fields()
        current["results"][0]["operation_metrics"] = operation_metrics_report_fields()
        current["results"][0]["operation_metrics"]["mystery"] = 1
        with self.assertRaisesRegex(
            perf_compare.ComparisonInputError,
            "operation_metrics keys mismatch.*mystery",
        ):
            perf_compare.compare_reports(
                baseline, current, self.operation_metrics_policy()
            )

        for field, value, pattern in (
            ("sample_count", 4, "sample_count=.*elapsed_ns.samples length"),
            ("alignment", "other.samples", "alignment must be"),
        ):
            with self.subTest(envelope_field=field):
                baseline = report()
                current = report(revision="current")
                baseline["results"][0]["operation_metrics"] = (
                    operation_metrics_report_fields()
                )
                current["results"][0]["operation_metrics"] = (
                    operation_metrics_report_fields()
                )
                current["results"][0]["operation_metrics"][field] = value
                with self.assertRaisesRegex(
                    perf_compare.ComparisonInputError, pattern
                ):
                    perf_compare.compare_reports(
                        baseline, current, self.operation_metrics_policy()
                    )

        for field, value, pattern in (
            ("sample_indices", [0, 0, 2, 3, 4], "sample_indices must be unique"),
            ("sample_indices", [0, 1, 2], "sample_indices has 3 samples"),
            (
                "sample_indices",
                [0, 1, 2, 3, 5],
                "sample_indices must be a complete permutation",
            ),
        ):
            with self.subTest(envelope_field=field, value=value):
                baseline = report()
                current = report(revision="current")
                baseline["results"][0]["operation_metrics"] = (
                    operation_metrics_report_fields()
                )
                current["results"][0]["operation_metrics"] = (
                    operation_metrics_report_fields()
                )
                current["results"][0]["operation_metrics"][field] = value
                with self.assertRaisesRegex(
                    perf_compare.ComparisonInputError, pattern
                ):
                    perf_compare.compare_reports(
                        baseline, current, self.operation_metrics_policy()
                    )

        baseline = report()
        current = report(revision="current")
        baseline["results"][0]["operation_metrics"] = operation_metrics_report_fields()
        current["results"][0]["operation_metrics"] = operation_metrics_report_fields()
        current["results"][0]["operation_metrics"]["sample_indices"] = [
            1,
            0,
            2,
            3,
            4,
        ]
        with self.assertRaisesRegex(
            perf_compare.ComparisonInputError,
            "increase across tied elapsed samples",
        ):
            perf_compare.compare_reports(
                baseline, current, self.operation_metrics_policy()
            )

        baseline = descriptive_parallel_report()
        current = descriptive_parallel_report(revision="current")
        for report_value in (baseline, current):
            sample_order = report_value["results"][0]["elapsed_ns"]["sample_order"]
            report_value["results"][0]["operation_metrics"] = (
                filesystem_xlsx_operation_metrics_report_fields(sample_order)
            )
        current["results"][0]["operation_metrics"]["sample_indices"] = list(
            range(5)
        )
        with self.assertRaisesRegex(
            perf_compare.ComparisonInputError,
            "sample_indices must match elapsed_ns.sample_order",
        ):
            perf_compare.compare_reports(
                baseline, current, self.operation_metrics_policy()
            )

        baseline = report()
        current = report(revision="current")
        baseline["results"][0]["operation_metrics"] = operation_metrics_report_fields()
        current["results"][0]["operation_metrics"] = operation_metrics_report_fields()
        current["results"][0]["operation_metrics"]["source"][
            "logical_read_pattern"
        ]["values"] = ["bogus"] * 5
        with self.assertRaisesRegex(
            perf_compare.ComparisonInputError, "logical_read_pattern.*one of"
        ):
            perf_compare.compare_reports(
                baseline, current, self.operation_metrics_policy()
            )

    def test_p50_and_p95_regressions_are_reported(self):
        result = perf_compare.compare_reports(report(), report(120, "current"), policy())
        self.assertEqual(result["status"], "regression")
        metrics = {item["metric"] for item in result["regressions"]}
        self.assertEqual(metrics, {"elapsed_ns.p50", "elapsed_ns.p95", "elapsed_ns.p99"})

    def test_optional_resource_regression_is_reported(self):
        current = report(revision="current")
        current["results"][0]["metrics"]["allocation_calls"] = [1100] * 5
        result = perf_compare.compare_reports(report(), current, policy())
        self.assertEqual(result["status"], "regression")
        regression = result["regressions"][0]
        self.assertEqual(regression["metric_class"], "allocation")
        self.assertEqual(regression["metric"], "metrics.allocation_calls")

    def test_allocator_instrumentation_withholds_elapsed_latency_claims(self):
        comparison_policy = allocator_policy_fixture()
        baseline = allocator_report()
        current = allocator_report(revision="current")
        for item in current["results"]:
            item["operation_metrics"]["allocation"]["allocation_calls"][
                "values"
            ] = [110] * 5
        for sample in current["filesystem_evidence"][0]["samples"]:
            sample["allocation_metrics"]["allocation_calls"] = 110
        result = perf_compare.compare_reports(baseline, current, comparison_policy)
        self.assertEqual(result["status"], "regression")
        self.assertEqual(result["summary"]["latency_claims"], "withheld_instrumentation")
        self.assertEqual(result["summary"]["latency_compared_results"], 0)
        self.assertFalse(
            any(item["metric_class"] == "latency" for item in result["comparisons"])
        )
        self.assertEqual(
            {item["metric"] for item in result["regressions"]},
            {"operation_metrics.allocation.allocation_calls.values"},
        )

    def test_allocator_filesystem_policy_requires_measured_vectors(self):
        comparison_policy = allocator_policy_fixture()
        baseline = allocator_report()
        current = allocator_report(revision="current")
        self.assertEqual(
            perf_compare.report_result_key_manifest_sha256(
                baseline,
                2,
                result_key_fields=("case", "corpus", "cache_state"),
            ),
            comparison_policy["expected_result_keys_sha256"],
        )
        result = perf_compare.compare_reports(baseline, current, comparison_policy)
        self.assertEqual(result["status"], "pass")
        self.assertEqual(result["summary"]["latency_claims"], "withheld_instrumentation")
        self.assertEqual(result["summary"]["compared_metrics"], 20)

        missing_raw = allocator_report(revision="current")
        del missing_raw["filesystem_evidence"][0]["samples"][0][
            "allocation_metrics"
        ]["allocated_bytes"]
        with self.assertRaisesRegex(
            perf_compare.ComparisonInputError,
            "raw allocation_metrics.*every numeric allocator field",
        ):
            perf_compare.compare_reports(baseline, missing_raw, comparison_policy)

        mismatched_raw = allocator_report(revision="current")
        mismatched_raw["filesystem_evidence"][0]["samples"][0][
            "allocation_metrics"
        ]["allocation_calls"] += 1
        with self.assertRaisesRegex(
            perf_compare.ComparisonInputError,
            "raw allocator allocation_calls.*does not match",
        ):
            perf_compare.compare_reports(baseline, mismatched_raw, comparison_policy)

        missing = allocator_report(revision="current")
        del missing["results"][0]["operation_metrics"]["allocation"][
            "live_bytes_after"
        ]["values"]
        with self.assertRaisesRegex(
            perf_compare.ComparisonInputError, "invalid schema|values"
        ):
            perf_compare.compare_reports(baseline, missing, comparison_policy)

        unavailable = allocator_report(revision="current")
        unavailable["results"][0]["operation_metrics"]["allocation"]["status"] = (
            "unavailable"
        )
        with self.assertRaisesRegex(
            perf_compare.ComparisonInputError, "status must be 'measured'"
        ):
            perf_compare.compare_reports(baseline, unavailable, comparison_policy)

        mismatched = allocator_report(revision="current")
        mismatched["results"][1]["operation_metrics"]["allocation"][
            "scope"
        ] = "other_scope"
        with self.assertRaisesRegex(
            perf_compare.ComparisonInputError, "scope must be"
        ):
            perf_compare.compare_reports(baseline, mismatched, comparison_policy)

    def test_allocator_policy_rejects_bad_cardinality_and_non_filesystem_rows(self):
        comparison_policy = allocator_policy_fixture()
        baseline = allocator_report()
        current = allocator_report(revision="current")
        current["results"][0]["operation_metrics"]["allocation"][
            "allocation_calls"
        ]["values"] = []
        with self.assertRaisesRegex(
            perf_compare.ComparisonInputError, "must be non-empty"
        ):
            perf_compare.compare_reports(baseline, current, comparison_policy)

        current = allocator_report(revision="current")
        current["filesystem_evidence"][0]["samples"].pop()
        with self.assertRaisesRegex(
            perf_compare.ComparisonInputError, "cardinality"
        ):
            perf_compare.compare_reports(baseline, current, comparison_policy)

        current = allocator_report(revision="current")
        current["filesystem_evidence"][0]["cache_states"] = ["warm"]
        with self.assertRaisesRegex(
            perf_compare.ComparisonInputError, "cache_states"
        ):
            perf_compare.compare_reports(baseline, current, comparison_policy)

        current = allocator_report(revision="current")
        current["results"][0]["case"] = "opc_open"
        current["filesystem_evidence"][0]["case"] = "opc_open"
        with self.assertRaisesRegex(
            perf_compare.ComparisonInputError, "filesystem case"
        ):
            perf_compare.compare_reports(baseline, current, comparison_policy)

    def test_metric_walker_treats_status_and_scope_as_metadata(self):
        comparison_policy = policy()
        comparison_policy["metric_classes"] = [
            {
                "name": "broad",
                "max_regression_percent": 5.0,
                "path_globs": ["**/*"],
                "presence": "optional",
            }
        ]
        baseline = report()
        current = report(revision="current")
        for item in (baseline, current):
            item["results"][0]["metrics"] = {
                "status": "measured",
                "scope": "operation_global_system_allocator",
                "sample_count": 5,
                "measured": [7] * 5,
            }
        result = perf_compare.compare_reports(baseline, current, comparison_policy)
        self.assertEqual(result["status"], "pass")
        metrics = {item["metric"] for item in result["comparisons"]}
        self.assertIn("metrics.measured", metrics)
        self.assertNotIn("metrics.status", metrics)
        self.assertNotIn("metrics.scope", metrics)

    def test_zero_baseline_metric_fails_on_nonzero_current(self):
        baseline = report()
        current = report(revision="current")
        baseline["results"][0]["source"]["read_calls"] = [0] * 5
        current["results"][0]["source"]["read_calls"] = [1] * 5
        result = perf_compare.compare_reports(baseline, current, policy())
        regression = next(
            item for item in result["regressions"] if item["metric"] == "source.read_calls"
        )
        self.assertTrue(regression["delta_is_infinite"])
        self.assertIsNone(regression["delta_percent"])

    def test_schema_tool_and_build_identity_mismatches_fail_closed(self):
        mutations = {
            "schema": lambda item: item.__setitem__("schema_version", 2),
            "tool": lambda item: item["tool"].__setitem__("version", "9.9.9"),
            "build": lambda item: item["environment"].__setitem__(
                "allocator", "jemalloc"
            ),
            "dirty": lambda item: item["environment"].__setitem__(
                "git_worktree_dirty", True
            ),
        }
        for name, mutate in mutations.items():
            with self.subTest(name=name):
                current = report(revision="current")
                mutate(current)
                with self.assertRaises(perf_compare.ComparisonInputError):
                    perf_compare.compare_reports(report(), current, policy())

    def test_case_and_corpus_key_mismatches_fail_closed(self):
        current = report(revision="current", corpus_sha="changed")
        with self.assertRaisesRegex(
            perf_compare.ComparisonInputError, "case/corpus key mismatch"
        ):
            perf_compare.compare_reports(report(), current, policy())

    def test_cache_state_is_part_of_result_identity(self):
        baseline = report()
        current = report(revision="current")
        corpus_identity = json.dumps(
            baseline["results"][0]["corpus"],
            sort_keys=True,
            separators=(",", ":"),
            allow_nan=False,
        )
        comparison_policy = policy()
        comparison_policy["expected_result_keys_sha256"] = (
            perf_compare.result_key_manifest_sha256(
                [("opc_open", corpus_identity, "warm")]
            )
        )
        baseline["results"][0]["cache_state"] = "warm"
        current["results"][0]["cache_state"] = "warm"
        result = perf_compare.compare_reports(baseline, current, comparison_policy)
        self.assertEqual(result["status"], "pass")
        self.assertTrue(
            all(item["cache_state"] == "warm" for item in result["comparisons"])
        )

        current["results"][0]["cache_state"] = "cold-requested"
        with self.assertRaisesRegex(
            perf_compare.ComparisonInputError, "case/corpus key mismatch"
        ):
            perf_compare.compare_reports(baseline, current, comparison_policy)

    def test_same_sided_corpus_replacement_fails_policy_manifest(self):
        baseline = report(corpus_sha="changed")
        current = report(revision="current", corpus_sha="changed")
        with self.assertRaisesRegex(perf_compare.ComparisonInputError, "manifest digest"):
            perf_compare.compare_reports(baseline, current, policy())
        current = report(revision="current")
        current["results"][0]["case"] = "other_case"
        with self.assertRaisesRegex(
            perf_compare.ComparisonInputError, "case/corpus key mismatch"
        ):
            perf_compare.compare_reports(report(), current, policy())

    def test_duplicate_and_missing_result_keys_fail_closed(self):
        duplicate_policy = policy()
        duplicate_policy["expected_result_count"] = 2
        baseline = report()
        current = report(revision="current")
        baseline["results"].append(copy.deepcopy(baseline["results"][0]))
        current["results"].append(copy.deepcopy(current["results"][0]))
        with self.assertRaisesRegex(perf_compare.ComparisonInputError, "duplicate"):
            perf_compare.compare_reports(baseline, current, duplicate_policy)

        current = report(revision="current")
        current["results"] = []
        with self.assertRaisesRegex(perf_compare.ComparisonInputError, "expected 1"):
            perf_compare.compare_reports(report(), current, policy())

    def test_extra_same_sided_case_fails_exact_policy_set(self):
        comparison_policy = policy()
        comparison_policy["expected_result_count"] = 2
        baseline = report()
        current = report(revision="current")
        for item in (baseline, current):
            rogue = copy.deepcopy(item["results"][0])
            rogue["case"] = "rogue_case"
            rogue["corpus"]["archive_sha256"] = "rogue"
            item["results"].append(rogue)
        with self.assertRaisesRegex(perf_compare.ComparisonInputError, "case set"):
            perf_compare.compare_reports(baseline, current, comparison_policy)

    def test_insufficient_samples_and_invalid_numbers_fail_closed(self):
        current = report(revision="current")
        current["results"][0]["elapsed_ns"]["samples"] = [100, 100]
        with self.assertRaisesRegex(perf_compare.ComparisonInputError, "minimum is 5"):
            perf_compare.compare_reports(report(), current, policy())

        current = report(revision="current")
        current["results"][0]["metrics"]["allocation_calls"] = [1, 2]
        with self.assertRaisesRegex(perf_compare.ComparisonInputError, "minimum is 5"):
            perf_compare.compare_reports(report(), current, policy())

        for invalid in (math.nan, math.inf, -1):
            with self.subTest(invalid=invalid):
                current = report(revision="current")
                current["results"][0]["elapsed_ns"]["samples"][0] = invalid
                with self.assertRaises(perf_compare.ComparisonInputError):
                    perf_compare.compare_reports(report(), current, policy())

        for invalid in ([], "bad", [1, "bad", 3, 4, 5]):
            with self.subTest(optional_metric=invalid):
                baseline = report()
                current = report(revision="current")
                baseline["results"][0]["source"]["read_calls"] = invalid
                current["results"][0]["source"]["read_calls"] = invalid
                with self.assertRaises(perf_compare.ComparisonInputError):
                    perf_compare.compare_reports(baseline, current, policy())

        current = report(revision="current")
        current["results"][0]["elapsed_ns"]["samples"][0] = 10**10000
        with self.assertRaisesRegex(perf_compare.ComparisonInputError, "finite range"):
            perf_compare.compare_reports(report(), current, policy())

    def test_reported_percentile_must_agree_with_samples(self):
        current = report(revision="current")
        current["results"][0]["elapsed_ns"]["p95"] = 999
        with self.assertRaisesRegex(perf_compare.ComparisonInputError, "disagrees"):
            perf_compare.compare_reports(report(), current, policy())

    def test_p99_is_required_and_must_agree_with_samples(self):
        current = report(revision="current")
        del current["results"][0]["elapsed_ns"]["p99"]
        with self.assertRaisesRegex(perf_compare.ComparisonInputError, "p99"):
            perf_compare.compare_reports(report(), current, policy())

        current = report(revision="current")
        current["results"][0]["elapsed_ns"]["p99"] = 999
        with self.assertRaisesRegex(perf_compare.ComparisonInputError, "p99.*disagrees"):
            perf_compare.compare_reports(report(), current, policy())

    def test_p99_rejects_nonfinite_and_overflow_values(self):
        for invalid in (math.nan, math.inf, -math.inf, 10**10000):
            with self.subTest(invalid=invalid):
                current = report(revision="current")
                current["results"][0]["elapsed_ns"]["p99"] = invalid
                with self.assertRaises(perf_compare.ComparisonInputError):
                    perf_compare.compare_reports(report(), current, policy())

    def test_reported_percentiles_must_be_non_decreasing(self):
        current = report(revision="current")
        elapsed = current["results"][0]["elapsed_ns"]
        elapsed["p50"] = 100.4
        elapsed["p95"] = 100.3
        elapsed["p99"] = 100.2
        with self.assertRaisesRegex(
            perf_compare.ComparisonInputError, "reported percentiles.*non-decreasing"
        ):
            perf_compare.compare_reports(report(), current, policy())

    def test_policy_and_report_schema_versions_are_strictly_typed(self):
        for invalid in (True, 1.0):
            with self.subTest(policy_schema=invalid):
                comparison_policy = policy()
                comparison_policy["schema_version"] = invalid
                with self.assertRaisesRegex(
                    perf_compare.ComparisonInputError,
                    "policy.schema_version must be an integer",
                ):
                    perf_compare.compare_reports(
                        report(), report(revision="current"), comparison_policy
                    )

            with self.subTest(report_schema=invalid):
                current = report(revision="current")
                current["schema_version"] = invalid
                with self.assertRaisesRegex(
                    perf_compare.ComparisonInputError,
                    "current.schema_version must be an integer",
                ):
                    perf_compare.compare_reports(report(), current, policy())

        legacy_policy = policy()
        legacy_policy["schema_version"] = 1
        with self.assertRaisesRegex(
            perf_compare.ComparisonInputError, "unsupported policy.schema_version"
        ):
            perf_compare.compare_reports(
                report(), report(revision="current"), legacy_policy
            )

    def test_required_metric_class_must_be_present_on_both_sides(self):
        comparison_policy = policy()
        comparison_policy["metric_classes"].append(
            {
                "name": "required_work",
                "max_regression_percent": 5.0,
                "path_globs": ["metrics/required_work"],
                "presence": "required",
            }
        )
        baseline = report()
        current = report(revision="current")
        baseline["results"][0]["metrics"]["required_work"] = [7] * 5
        current["results"][0]["metrics"]["required_work"] = [7] * 5
        result = perf_compare.compare_reports(baseline, current, comparison_policy)
        self.assertEqual(result["status"], "pass")
        self.assertEqual(result["summary"]["compared_metrics"], 8)

        del current["results"][0]["metrics"]["required_work"]
        with self.assertRaisesRegex(
            perf_compare.ComparisonInputError, "missing required metrics"
        ):
            perf_compare.compare_reports(baseline, current, comparison_policy)

    def test_required_metric_vector_is_validated_and_short_vectors_fail_closed(self):
        comparison_policy = policy()
        comparison_policy["metric_classes"].append(
            {
                "name": "required_work",
                "max_regression_percent": 5.0,
                "path_globs": ["metrics/required_work"],
                "presence": "required",
            }
        )
        baseline = report()
        current = report(revision="current")
        for item in (baseline, current):
            item["results"][0]["metrics"]["required_work"] = [7] * 4
        with self.assertRaisesRegex(
            perf_compare.ComparisonInputError, "required metric.*minimum is 5"
        ):
            perf_compare.compare_reports(baseline, current, comparison_policy)

    def test_policy_schema_requires_explicit_metric_presence(self):
        comparison_policy = policy()
        comparison_policy["metric_classes"][0].pop("presence")
        with self.assertRaisesRegex(
            perf_compare.ComparisonInputError, "presence is required"
        ):
            perf_compare.compare_reports(
                report(), report(revision="current"), comparison_policy
            )

    def test_metric_presence_policy_is_strict(self):
        comparison_policy = policy()
        comparison_policy["metric_classes"][0]["presence"] = "sometimes"
        with self.assertRaisesRegex(
            perf_compare.ComparisonInputError, "presence must be 'required' or 'optional'"
        ):
            perf_compare.compare_reports(
                report(), report(revision="current"), comparison_policy
            )

    def test_optional_metric_presence_must_match(self):
        current = report(revision="current")
        del current["results"][0]["metrics"]["peak_rss_bytes"]
        with self.assertRaisesRegex(
            perf_compare.ComparisonInputError, "optional metric mismatch"
        ):
            perf_compare.compare_reports(report(), current, policy())

    def test_configuration_must_match(self):
        current = report(revision="current")
        current["configuration"]["samples_per_case"] = 6
        with self.assertRaisesRegex(perf_compare.ComparisonInputError, "configuration"):
            perf_compare.compare_reports(report(), current, policy())

        current = report(revision="current")
        current["configuration"]["cases"] = ["rogue_case"]
        with self.assertRaisesRegex(perf_compare.ComparisonInputError, "exact policy"):
            perf_compare.compare_reports(report(), current, policy())

        baseline = report()
        current = report(revision="current")
        baseline["configuration"]["execution_workers"] = [1]
        current["configuration"]["execution_workers"] = [1]
        with self.assertRaisesRegex(perf_compare.ComparisonInputError, "derived default"):
            perf_compare.compare_reports(baseline, current, policy())

        baseline = report()
        current = report(revision="current")
        current["configuration"]["opc_cache_lock_diagnostics"] = False
        self.assertEqual(
            perf_compare.compare_reports(baseline, current, policy())["status"],
            "pass",
        )
        current["configuration"]["opc_cache_lock_diagnostics"] = True
        with self.assertRaisesRegex(
            perf_compare.ComparisonInputError, "benchmark configuration mismatch"
        ):
            perf_compare.compare_reports(baseline, current, policy())

        baseline = report()
        current = report(revision="current")
        current["configuration"][
            "xlsx_cell_values_managed_planning_memory_headroom"
        ] = 64 * 1024
        with self.assertRaisesRegex(
            perf_compare.ComparisonInputError, "benchmark configuration mismatch"
        ):
            perf_compare.compare_reports(baseline, current, policy())

    def test_managed_xlsx_planning_allowance_must_be_unsigned(self):
        baseline, current, comparison_policy = managed_xlsx_reports()
        current["configuration"][
            "xlsx_cell_values_managed_planning_memory_headroom"
        ] = True
        with self.assertRaisesRegex(
            perf_compare.ComparisonInputError, "unsigned 64-bit integer scalar"
        ):
            perf_compare.compare_reports(baseline, current, comparison_policy)

    def test_managed_xlsx_planning_allowance_is_required_by_new_source_evidence(self):
        baseline, current, comparison_policy = managed_xlsx_reports()
        del current["configuration"][
            "xlsx_cell_values_managed_planning_memory_headroom"
        ]
        with self.assertRaisesRegex(
            perf_compare.ComparisonInputError, "required for new managed XLSX"
        ):
            perf_compare.compare_reports(baseline, current, comparison_policy)

    def test_managed_xlsx_planning_allowance_must_match_source_evidence(self):
        baseline, current, comparison_policy = managed_xlsx_reports()
        current["configuration"][
            "xlsx_cell_values_managed_planning_memory_headroom"
        ] += 1
        with self.assertRaisesRegex(
            perf_compare.ComparisonInputError, "does not match"
        ):
            perf_compare.compare_reports(baseline, current, comparison_policy)

    def test_managed_xlsx_resulting_memory_limit_must_include_planning_allowance(self):
        baseline, current, comparison_policy = managed_xlsx_reports()
        current["results"][0]["source"]["xlsx_cell_values"][
            "cache_budget_memory_limit"
        ] += 1
        with self.assertRaisesRegex(
            perf_compare.ComparisonInputError, "must equal payload_memory_limit"
        ):
            perf_compare.compare_reports(baseline, current, comparison_policy)

    def test_managed_xlsx_memory_tuple_is_validated(self):
        baseline, current, comparison_policy = managed_xlsx_reports()
        self.assertEqual(
            perf_compare.compare_reports(baseline, current, comparison_policy)["status"],
            "pass",
        )
        for field in (
            "payload_memory_limit",
            "publication_planning_memory_headroom",
            "cache_budget_memory_limit",
        ):
            with self.subTest(missing=field):
                malformed = copy.deepcopy(current)
                del malformed["results"][0]["source"]["xlsx_cell_values"][field]
                with self.assertRaises(perf_compare.ComparisonInputError):
                    perf_compare.compare_reports(baseline, malformed, comparison_policy)
            for value in (True, -1, 1.5, None, 2**64):
                with self.subTest(field=field, invalid=value):
                    malformed = copy.deepcopy(current)
                    malformed["results"][0]["source"]["xlsx_cell_values"][field] = value
                    with self.assertRaises(perf_compare.ComparisonInputError):
                        perf_compare.compare_reports(baseline, malformed, comparison_policy)

    def test_managed_xlsx_memory_tuple_rejects_sum_overflow(self):
        baseline, current, comparison_policy = managed_xlsx_reports()
        current["results"][0]["source"]["xlsx_cell_values"][
            "payload_memory_limit"
        ] = 2**64 - 1
        with self.assertRaisesRegex(perf_compare.ComparisonInputError, "overflows"):
            perf_compare.compare_reports(baseline, current, comparison_policy)

    def test_legacy_managed_xlsx_reports_remain_comparable(self):
        baseline, current, comparison_policy = legacy_managed_xlsx_reports()
        self.assertEqual(
            perf_compare.compare_reports(baseline, current, comparison_policy)["status"],
            "pass",
        )

    def test_unapproved_corpus_manifest_fails_closed(self):
        comparison_policy = policy()
        comparison_policy["expected_result_keys_sha256"] = None
        with self.assertRaisesRegex(perf_compare.ComparisonInputError, "no approved"):
            perf_compare.compare_reports(
                report(), report(revision="current"), comparison_policy
            )

    def test_build_identity_values_are_typed_and_nullable_only_by_policy(self):
        current = report(revision="current")
        current["environment"]["allocator"] = None
        with self.assertRaisesRegex(perf_compare.ComparisonInputError, "must not be null"):
            perf_compare.compare_reports(report(), current, policy())

        nullable_policy = policy()
        nullable_policy["nullable_build_identity_fields"] = ["rustflags"]
        baseline = report()
        current = report(revision="current")
        baseline["environment"]["rustflags"] = None
        current["environment"]["rustflags"] = None
        result = perf_compare.compare_reports(baseline, current, nullable_policy)
        self.assertEqual(result["status"], "pass")

    def test_reference_and_current_revisions_must_differ(self):
        with self.assertRaisesRegex(perf_compare.ComparisonInputError, "must differ"):
            perf_compare.compare_reports(report(), report(), policy())

    def test_cli_emits_machine_and_human_reports_and_exit_codes(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            policy_path = root / "policy.json"
            baseline_path = root / "baseline.json"
            current_path = root / "current.json"
            json_path = root / "comparison.json"
            text_path = root / "comparison.txt"
            policy_path.write_text(json.dumps(policy()), encoding="utf-8")
            baseline_path.write_text(json.dumps(report()), encoding="utf-8")
            current_path.write_text(
                json.dumps(report(104, "current")), encoding="utf-8"
            )
            args = [
                "--policy",
                str(policy_path),
                "--baseline",
                str(baseline_path),
                "--current",
                str(current_path),
                "--json-out",
                str(json_path),
                "--summary-out",
                str(text_path),
            ]
            self.assertEqual(perf_compare.main(args), 0)
            self.assertEqual(json.loads(json_path.read_text())["status"], "pass")
            self.assertIn("PASS:", text_path.read_text())

            current_path.write_text(
                json.dumps(report(120, "current")), encoding="utf-8"
            )
            self.assertEqual(perf_compare.main(args), 1)
            self.assertEqual(json.loads(json_path.read_text())["status"], "regression")

            current_path.write_text("{invalid", encoding="utf-8")
            self.assertEqual(perf_compare.main(args), 2)
            self.assertEqual(json.loads(json_path.read_text())["status"], "invalid")
            self.assertIn("INVALID:", text_path.read_text())


    def test_opc_zip_metrics_preserve_zeroes_and_omit_non_measured_values(self):
        baseline = report()
        current = report(revision="current")
        for item in (baseline["results"][0], current["results"][0]):
            item["elapsed_ns"]["sample_order"] = list(range(5))
            item["operation_metrics"] = opc_zip_operation_metrics_report_fields()
        add_opc_zip_sample_order(baseline["results"][0], current["results"][0])
        result = perf_compare.compare_reports(
            baseline, current, self.operation_metrics_policy()
        )
        self.assertEqual(result["status"], "pass")
        opc_zip = current["results"][0]["operation_metrics"]["opc_zip"]
        self.assertEqual(
            opc_zip["output_bytes_accepted"]["values"], [0] * 5
        )

        for status in ("not_applicable", "unavailable", "overflow"):
            with self.subTest(status=status):
                baseline = report()
                current = report(revision="current")
                for item in (baseline["results"][0], current["results"][0]):
                    item["elapsed_ns"]["sample_order"] = list(range(5))
                    item["operation_metrics"] = opc_zip_operation_metrics_report_fields(
                        status=status
                    )
                add_opc_zip_sample_order(baseline["results"][0], current["results"][0])
                result = perf_compare.compare_reports(
                    baseline, current, self.operation_metrics_policy()
                )
                self.assertEqual(result["status"], "pass")

    def test_opc_zip_metrics_reject_unknown_invalid_and_misaligned_fields(self):
        baseline = report()
        current = report(revision="current")
        baseline["results"][0]["operation_metrics"] = (
            opc_zip_operation_metrics_report_fields()
        )
        current["results"][0]["operation_metrics"] = (
            opc_zip_operation_metrics_report_fields()
        )
        add_opc_zip_sample_order(baseline["results"][0], current["results"][0])
        current["results"][0]["operation_metrics"]["opc_zip"]["mystery"] = 1
        with self.assertRaisesRegex(
            perf_compare.ComparisonInputError,
            "operation_metrics.opc_zip keys mismatch.*mystery",
        ):
            perf_compare.compare_reports(
                baseline, current, self.operation_metrics_policy()
            )

        baseline = report()
        current = report(revision="current")
        baseline["results"][0]["operation_metrics"] = (
            opc_zip_operation_metrics_report_fields()
        )
        current["results"][0]["operation_metrics"] = (
            opc_zip_operation_metrics_report_fields()
        )
        add_opc_zip_sample_order(baseline["results"][0], current["results"][0])
        del current["results"][0]["operation_metrics"]["opc_zip"][
            "output_bytes_accepted"
        ]
        with self.assertRaisesRegex(
            perf_compare.ComparisonInputError,
            "operation_metrics.opc_zip keys mismatch.*output_bytes_accepted",
        ):
            perf_compare.compare_reports(
                baseline, current, self.operation_metrics_policy()
            )

        invalid_values = (
            [0] * 4,
            [0, 0, False, 0, 0],
            [0, 0, -1, 0, 0],
            [0, 0, "bad", 0, 0],
        )
        for invalid in invalid_values:
            with self.subTest(invalid=invalid):
                baseline = report()
                current = report(revision="current")
                baseline["results"][0]["operation_metrics"] = (
                    opc_zip_operation_metrics_report_fields()
                )
                current["results"][0]["operation_metrics"] = (
                    opc_zip_operation_metrics_report_fields()
                )
                add_opc_zip_sample_order(
                    baseline["results"][0], current["results"][0]
                )
                current["results"][0]["operation_metrics"]["opc_zip"][
                    "output_bytes_accepted"
                ]["values"] = invalid
                with self.assertRaises(perf_compare.ComparisonInputError):
                    perf_compare.compare_reports(
                        baseline, current, self.operation_metrics_policy()
                    )

        baseline = report()
        current = report(revision="current")
        baseline["results"][0]["operation_metrics"] = (
            opc_zip_operation_metrics_report_fields()
        )
        current["results"][0]["operation_metrics"] = (
            opc_zip_operation_metrics_report_fields()
        )
        add_opc_zip_sample_order(baseline["results"][0], current["results"][0])
        del current["results"][0]["operation_metrics"]["opc_zip"][
            "output_bytes_accepted"
        ]["values"]
        with self.assertRaisesRegex(
            perf_compare.ComparisonInputError,
            "opc_zip.output_bytes_accepted.values is required",
        ):
            perf_compare.compare_reports(
                baseline, current, self.operation_metrics_policy()
            )

        non_measured = opc_zip_operation_metrics_report_fields(status="unavailable")
        non_measured["opc_zip"]["output_bytes_accepted"]["values"] = [0] * 5
        baseline = report()
        current = report(revision="current")
        baseline["results"][0]["operation_metrics"] = (
            opc_zip_operation_metrics_report_fields(status="unavailable")
        )
        current["results"][0]["operation_metrics"] = non_measured
        add_opc_zip_sample_order(baseline["results"][0], current["results"][0])
        with self.assertRaisesRegex(
            perf_compare.ComparisonInputError,
            "opc_zip.output_bytes_accepted.values must be omitted",
        ):
            perf_compare.compare_reports(
                baseline, current, self.operation_metrics_policy()
            )

        baseline = report()
        current = report(revision="current")
        baseline["results"][0]["operation_metrics"] = (
            opc_zip_operation_metrics_report_fields()
        )
        current["results"][0]["operation_metrics"] = (
            opc_zip_operation_metrics_report_fields()
        )
        add_opc_zip_sample_order(baseline["results"][0], current["results"][0])
        current["results"][0]["operation_metrics"]["opc_zip"]["scope"] = (
            "other_scope"
        )
        with self.assertRaisesRegex(
            perf_compare.ComparisonInputError, "opc_zip.scope must be"
        ):
            perf_compare.compare_reports(
                baseline, current, self.operation_metrics_policy()
            )


    def test_opc_zip_claim_alignment_and_u64_validation_fail_closed(self):
        baseline = report()
        current = report(revision="current")
        for item in (baseline["results"][0], current["results"][0]):
            item["elapsed_ns"]["sample_order"] = list(range(5))
            item["operation_metrics"] = operation_metrics_report_fields()
            item["operation_metrics"]["latency_claim"] = (
                perf_compare.OPC_ZIP_EVIDENCE_ONLY_LATENCY_CLAIM
            )
        with self.assertRaisesRegex(
            perf_compare.ComparisonInputError,
            "latency_claim and operation_metrics.opc_zip must be present together",
        ):
            perf_compare.compare_reports(
                baseline, current, self.operation_metrics_policy()
            )

        baseline = report()
        current = report(revision="current")
        for item in (baseline["results"][0], current["results"][0]):
            item["elapsed_ns"]["sample_order"] = list(range(5))
            item["operation_metrics"] = opc_zip_operation_metrics_report_fields()
            item["operation_metrics"]["latency_claim"] = (
                perf_compare.EVIDENCE_ONLY_LATENCY_CLAIM
            )
        with self.assertRaisesRegex(
            perf_compare.ComparisonInputError,
            "latency_claim and operation_metrics.opc_zip must be present together",
        ):
            perf_compare.compare_reports(
                baseline, current, self.operation_metrics_policy()
            )

        baseline = report()
        current = report(revision="current")
        for item in (baseline["results"][0], current["results"][0]):
            item["elapsed_ns"]["sample_order"] = list(range(5))
            item["operation_metrics"] = opc_zip_operation_metrics_report_fields()
        current["results"][0]["elapsed_ns"]["sample_order"] = [1, 0, 2, 3, 4]
        with self.assertRaisesRegex(
            perf_compare.ComparisonInputError,
            "opc_zip sample_indices must match elapsed_ns.sample_order",
        ):
            perf_compare.compare_reports(
                baseline, current, self.operation_metrics_policy()
            )

        for invalid in ([0, 0, 1 << 64, 0, 0], [0, 0, 1.0, 0, 0]):
            with self.subTest(invalid=invalid):
                baseline = report()
                current = report(revision="current")
                for item in (baseline["results"][0], current["results"][0]):
                    item["elapsed_ns"]["sample_order"] = list(range(5))
                    item["operation_metrics"] = opc_zip_operation_metrics_report_fields()
                current["results"][0]["operation_metrics"]["opc_zip"][
                    "output_bytes_accepted"
                ]["values"] = invalid
                with self.assertRaisesRegex(
                    perf_compare.ComparisonInputError,
                    "must be a non-negative integer",
                ):
                    perf_compare.compare_reports(
                        baseline, current, self.operation_metrics_policy()
                    )

        baseline = report()
        current = report(revision="current")
        for item in (baseline["results"][0], current["results"][0]):
            item["elapsed_ns"]["sample_order"] = list(range(5))
            item["operation_metrics"] = opc_zip_operation_metrics_report_fields()
        current["results"][0]["operation_metrics"]["opc_zip"]["status"] = (
            "unavailable"
        )
        with self.assertRaisesRegex(
            perf_compare.ComparisonInputError,
            "opc_zip.status does not match",
        ):
            perf_compare.compare_reports(
                baseline, current, self.operation_metrics_policy()
            )


if __name__ == "__main__":
    unittest.main()
