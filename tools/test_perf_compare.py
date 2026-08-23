import copy
import json
import math
import tempfile
import unittest
from pathlib import Path

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
    result["filesystem_evidence"] = [
        {
            "case": "opc_file_eager_open",
            "sample_count": 5,
            "cache_states": ["warm", "cold-requested"],
            "samples": [
                {
                    "sample_index": sample_index,
                    "cache_state": cache_state,
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


if __name__ == "__main__":
    unittest.main()
