import copy
import json
import math
import tempfile
import unittest
from pathlib import Path

from tools import perf_abba_summary
from tools import perf_opc_overlay_abba_summary as summary


def _statistics(samples, sample_order):
    ordered = list(samples)
    if ordered != sorted(ordered):
        raise AssertionError("test fixture expects sorted elapsed samples")
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
    margin = (
        perf_abba_summary._student_t_critical_95(len(ordered) - 1)
        * standard_deviation
        / math.sqrt(len(ordered))
    )
    left = ordered[(len(ordered) - 1) // 2]
    right = ordered[len(ordered) // 2]

    def nearest_rank(percentile):
        index = ((percentile * len(ordered) + 99) // 100) - 1
        return ordered[min(index, len(ordered) - 1)]

    return {
        "unit": "ns",
        "sample_order": list(sample_order),
        "min": ordered[0],
        "p50": left // 2 + right // 2 + (left % 2 + right % 2) // 2,
        "p95": nearest_rank(95),
        "p99": nearest_rank(99),
        "max": ordered[-1],
        "mean": mean,
        "standard_deviation": standard_deviation,
        "confidence_interval_95": {
            "method": "two-sided Student's t interval for the mean",
            "lower": max(mean - margin, 0.0),
            "upper": mean + margin,
        },
        "samples": ordered,
    }


def _corpus(shape, count):
    expected = summary.EXPECTED_CORPORA[shape]
    return {
        field: value
        for field, value in {
            **expected,
            "name": f"{expected['name_prefix']}{count}",
        }.items()
        if field != "name_prefix"
    }


def _metric_vector(status, scope, sample_count, values=None):
    result = {"status": status, "scope": scope}
    if values is not None:
        result["values"] = list(values)
    return result


def _operation_metrics(sample_count, sample_order, sink):
    absent_source = {
        "status": "not_applicable",
        "counter_scope": "not_applicable_in_process_sink",
        "logical_read_calls": _metric_vector(
            "not_applicable", summary.SOURCE_SCOPE, sample_count
        ),
        "logical_read_requested_bytes": _metric_vector(
            "not_applicable", summary.SOURCE_SCOPE, sample_count
        ),
        "logical_read_returned_bytes": _metric_vector(
            "not_applicable", summary.SOURCE_SCOPE, sample_count
        ),
        "logical_read_largest_requested_bytes": _metric_vector(
            "not_applicable", summary.SOURCE_SCOPE, sample_count
        ),
        "logical_read_largest_returned_bytes": _metric_vector(
            "not_applicable", summary.SOURCE_SCOPE, sample_count
        ),
        "logical_read_pattern": _metric_vector(
            "not_applicable", summary.SOURCE_PATTERN_SCOPE, sample_count
        ),
        "compressed_bytes": _metric_vector(
            "not_applicable", summary.SOURCE_COMPRESSED_SCOPE, sample_count
        ),
        "decompressed_bytes": _metric_vector(
            "not_applicable", summary.SOURCE_DECOMPRESSED_SCOPE, sample_count
        ),
        "recompressed_bytes": _metric_vector(
            "not_applicable", summary.SOURCE_RECOMPRESSED_SCOPE, sample_count
        ),
        "max_concurrent_reads": _metric_vector(
            "not_applicable", summary.SOURCE_SCOPE, sample_count
        ),
    }
    process_scopes = {
        "user_cpu_ticks": summary.IN_PROCESS_PROCESS_SCOPE,
        "system_cpu_ticks": summary.IN_PROCESS_PROCESS_SCOPE,
        "clock_ticks_per_second": summary.IN_PROCESS_PROCESS_SCOPE,
        "minor_faults": summary.IN_PROCESS_PROCESS_SCOPE,
        "major_faults": summary.IN_PROCESS_PROCESS_SCOPE,
        "voluntary_context_switches": summary.IN_PROCESS_PROCESS_SCOPE,
        "nonvoluntary_context_switches": summary.IN_PROCESS_PROCESS_SCOPE,
        "rss_delta_bytes": summary.IN_PROCESS_RSS_SCOPE,
        "peak_rss_bytes": summary.HWM_SCOPE,
        "rchar": summary.IN_PROCESS_PROCESS_SCOPE,
        "wchar": summary.IN_PROCESS_PROCESS_SCOPE,
        "read_bytes": summary.IN_PROCESS_PROCESS_SCOPE,
        "write_bytes": summary.IN_PROCESS_PROCESS_SCOPE,
        "cancelled_write_bytes": summary.IN_PROCESS_PROCESS_SCOPE,
        "syscr": summary.IN_PROCESS_PROCESS_SCOPE,
        "syscw": summary.IN_PROCESS_PROCESS_SCOPE,
    }
    process = {"status": "unavailable"}
    process.update(
        {
            field: _metric_vector("unavailable", scope, sample_count)
            for field, scope in process_scopes.items()
        }
    )
    operation_sink = {
        "status": "not_applicable",
        "output_bytes": _metric_vector(
            "not_applicable", summary.OUTPUT_BYTES_SCOPE, sample_count
        ),
        "write_status": "measured",
        "accepted_bytes": _metric_vector(
            "measured",
            summary.SINK_ACCEPTED_BYTES_SCOPE,
            sample_count,
            [sink["accepted_bytes"]] * sample_count,
        ),
        "write_calls": _metric_vector(
            "measured",
            summary.SINK_WRITE_CALLS_SCOPE,
            sample_count,
            [sink["write_calls"]] * sample_count,
        ),
        "largest_write": _metric_vector(
            "measured",
            summary.SINK_LARGEST_WRITE_SCOPE,
            sample_count,
            [sink["largest_write"]] * sample_count,
        ),
    }
    operation_sink["write_size_buckets"] = {
        "status": "measured",
        **{
            field: _metric_vector(
                "measured",
                summary.SINK_BUCKET_SCOPE,
                sample_count,
                [sink["write_size_buckets"][field]] * sample_count,
            )
            for field in summary.SINK_BUCKET_KEYS
        },
    }
    absent_publication = {
        "status": "not_applicable",
        "changed_spans": _metric_vector(
            "not_applicable", "logical_publication_counter", sample_count
        ),
        "published_bytes": _metric_vector(
            "not_applicable", "logical_publication_counter", sample_count
        ),
    }
    absent_materialization = {
        "status": "not_applicable",
        "opc_parts": _metric_vector(
            "not_applicable", "logical_materialization_counter", sample_count
        ),
    }
    absent_phases = {
        "status": "not_applicable",
        **{
            phase: {
                "elapsed_ns": _metric_vector(
                    "not_applicable", summary.CFB_PHASE_ELAPSED_SCOPE, sample_count
                ),
                "logical_read_calls": _metric_vector(
                    "not_applicable", summary.CFB_PHASE_SOURCE_SCOPE, sample_count
                ),
                "logical_read_requested_bytes": _metric_vector(
                    "not_applicable", summary.CFB_PHASE_SOURCE_SCOPE, sample_count
                ),
                "logical_read_returned_bytes": _metric_vector(
                    "not_applicable", summary.CFB_PHASE_SOURCE_SCOPE, sample_count
                ),
            }
            for phase in ("open", "plan", "atomic_publication")
        },
    }
    allocation = {
        "status": "unavailable",
        "scope": summary.ALLOCATION_SCOPE,
        **{
            field: _metric_vector(
                "unavailable", summary.ALLOCATION_SCOPE, sample_count
            )
            for field in summary.ALLOCATION_VECTOR_FIELDS
        },
    }
    return {
        "sample_count": sample_count,
        "sample_indices": list(sample_order),
        "alignment": summary.OPERATION_ALIGNMENT,
        "latency_claim": summary.OPERATION_LATENCY_CLAIM,
        "source": absent_source,
        "process": process,
        "sink": operation_sink,
        "publication": absent_publication,
        "materialization": absent_materialization,
        "cfb_phases": absent_phases,
        "allocation": allocation,
    }


def _source_and_sink(shape, count, role, sample_count):
    expected = summary.EXPECTED_CORPORA[shape]
    role_base = {"a1": 1000, "b1": 800, "b2": 810, "a2": 1010}[role]
    publication_original = [
        role_base + count * 10 + index + (index % 3)
        for index in range(sample_count)
    ]
    preparation_original = [3 + (index % 2) for index in range(sample_count)]
    open_original = [5 + (index % 3) for index in range(sample_count)]
    planning_original = [2 + (index % 2) for index in range(sample_count)]
    elapsed_original = [
        publication_original[index]
        + preparation_original[index]
        + open_original[index]
        + planning_original[index]
        for index in range(sample_count)
    ]
    sample_order = sorted(
        range(sample_count), key=lambda index: (elapsed_original[index], index)
    )
    elapsed_sorted = [elapsed_original[index] for index in sample_order]

    def aligned(values):
        return [values[index] for index in sample_order]

    digest = expected["archive_sha256"]
    overlay = {
        "implementation": "SourceBackedPackage::write_part_overlays_to_stream",
        "timing_scope": summary.TIMING_SCOPE,
        "performance_claim": "none",
        "overlay_mode": "noop",
        "replacement_semantics": "non-empty equal-payload replacement plan; semantic no-op",
        "overlay_count": count,
        "source_shape": shape,
        "payload_kind": expected["payload_kind"],
        "source_bytes": expected["archive_bytes"],
        "source_sha256": digest,
        "expected_eager_sha256": digest,
        "source_cache_max_bytes": expected["uncompressed_payload_bytes"],
        "source_cache_max_entries": expected["entry_count"],
        "sink_max_bytes": 2 * expected["archive_bytes"] + 65_536,
        "sink_max_write": 65_536,
        "preparation_ns": aligned(preparation_original),
        "open_ns": aligned(open_original),
        "planning_ns": aligned(planning_original),
        "publication_ns": aligned(publication_original),
        "cache_before_publication_hits": [0] * sample_count,
        "cache_before_publication_cold_loads": [0] * sample_count,
        "cache_before_publication_retained_entries": [0] * sample_count,
        "cache_before_publication_retained_bytes": [0] * sample_count,
        "source_cache_after_publication_probe_hits": [0] * sample_count,
        "source_cache_after_publication_probe_cold_loads": [count] * sample_count,
        "source_cache_after_publication_probe_retained_entries": [count] * sample_count,
        "source_cache_after_publication_probe_retained_bytes": [
            count * expected["entry_bytes"]
        ]
        * sample_count,
        "reopened_output_cache_hits": [0] * sample_count,
        "reopened_output_cache_cold_loads": [0] * sample_count,
        "reopened_output_cache_retained_entries": [0] * sample_count,
        "reopened_output_cache_retained_bytes": [0] * sample_count,
        "observed_after_publication_source_read_calls": [count + 3] * sample_count,
        "observed_after_publication_source_read_bytes": [expected["archive_bytes"]] * sample_count,
        "observed_after_publication_ordinary_payload_read_calls": [count] * sample_count,
        "observed_after_publication_ordinary_payload_read_bytes": [
            count * expected["entry_bytes"]
        ]
        * sample_count,
        "expected_eager_semantic_verified": True,
        "raw_members_and_order_preservation_verified": True,
        "equal_payload_noop_source_verified": True,
        "observed_output_sha256": [digest] * sample_count,
    }
    source = {
        "read_calls": [count + 3] * sample_count,
        "read_bytes": [expected["archive_bytes"]] * sample_count,
        "ordinary_payload_read_calls": [count] * sample_count,
        "ordinary_payload_read_bytes": [count * expected["entry_bytes"]] * sample_count,
        "max_in_flight_reads": [1] * sample_count,
        "opc_source_overlay": overlay,
    }
    accepted_bytes = expected["archive_bytes"]
    write_calls = max(3, (accepted_bytes + 65_535) // 65_536)
    largest_write = 512 if accepted_bytes <= 16_384 else 65_536
    sink = {
        "accepted_bytes": accepted_bytes,
        "write_calls": write_calls,
        "largest_write": largest_write,
        "write_size_buckets": {
            "bytes_0": 0,
            "bytes_1_to_512": write_calls if largest_write == 512 else 0,
            "bytes_513_to_4096": 0,
            "bytes_4097_to_16384": 0,
            "bytes_16385_to_65536": write_calls if largest_write == 65_536 else 0,
            "bytes_over_65536": 0,
        },
    }
    return source, sink, elapsed_sorted, sample_order


def _report(revision, binary_digest, role, sample_count=summary.SAMPLE_COUNT):
    results = []
    for shape in summary.SHAPES:
        for count in summary.COUNTS:
            source, sink, elapsed_samples, sample_order = _source_and_sink(
                shape, count, role, sample_count
            )
            row = {
                "case": summary.CASE,
                "corpus": _corpus(shape, count),
                "elapsed_ns": _statistics(elapsed_samples, sample_order),
                "sink": sink,
                "source": source,
                "output_sha256": summary.EXPECTED_CORPORA[shape]["archive_sha256"],
                "operation_metrics": _operation_metrics(
                    sample_count, sample_order, sink
                ),
            }
            results.append(row)
    configuration = {
        "samples_per_case": sample_count,
        "warmup_iterations_per_case": summary.WARMUP_ITERATIONS,
        "filesystem_cache_states": ["warm", "cold-requested"],
        "filesystem_fresh_child_per_sample": True,
        "filesystem_process_isolated": True,
        "filesystem_root_selected": False,
        "cases": [summary.CASE],
        "corpus_shapes": ["tiny", "many-small", "few-large", "wide-root"],
        "payload_kinds": ["compressible", "incompressible"],
        "writer_shapes": ["tiny", "large", "payload-heavy"],
        "xlsx_shapes": ["tiny", "medium", "dense-wide"],
        "xlsb_shapes": ["tiny", "medium", "large", "sparse"],
        "xlsx_cell_crud_shapes": ["medium", "dense-sparse"],
        "xlsx_row_visibility_shapes": ["medium", "large"],
        "semantic_shapes": ["tiny", "medium", "large"],
        "rtf_variants": ["plain"],
        "range_simulation": {
            "fixed_latency_us": 100,
            "request_overhead_us": 25,
            "bandwidth_bytes_per_second": 52_428_800,
            "max_physical_range_bytes": 4096,
        },
        "execution_workers": [1],
    }
    parallel_cases = []
    for row in results:
        parallel_cases.append(
            {
                "case": row["case"],
                "corpus_sha256": row["corpus"]["archive_sha256"],
                "configured_worker_count": {
                    "status": "not_applicable",
                    "scope": "result.execution.worker_count",
                    "reason": "overlay operation has no explicit execution worker",
                },
                "observed_local_worker_count": {
                    "status": "not_applicable",
                    "scope": "result.source.opc_cache.worker_count_with_one_created_local_worker_team",
                    "reason": "overlay operation creates no local worker team",
                },
                "deterministic_task_count": {
                    "status": "not_applicable",
                    "scope": "result.execution.logical_tasks",
                    "reason": "overlay operation has no explicit execution context",
                },
                "deterministic_chunk_count": {
                    "status": "not_applicable",
                    "scope": "result.execution.deterministic_chunk_count",
                    "reason": "overlay operation has no execution chunks",
                },
                "lock_wait_ns": {
                    "status": "unavailable",
                    "scope": "lock_wait_ns",
                    "reason": "no exact instrumented lock boundary is present",
                },
            }
        )
    return {
        "schema_version": 1,
        "tool": {
            "name": "litchi-perf-baseline",
            "version": "0.1.0",
            "binary": "litchi-perf-baseline",
            "profile": "release",
            "target_os": "linux",
            "target_arch": "x86_64",
            "instrumentation": "none",
        },
        "binary_identity": {
            "path": "/tmp/litchi-perf-baseline",
            "binary_sha256": binary_digest,
            "binary_bytes": 1000,
            "mode_bits": 0o755,
            "executable": True,
            "profile": "release",
        },
        "environment": {
            "rustc_version": summary.RUSTC_VERSION,
            "git_revision": revision,
            "git_worktree_dirty": False,
            "logical_cpus_available": 32,
            "allocator": "Rust system allocator",
            "rustflags": None,
            "cargo_build_target": None,
            "perf_event_paranoid": "1",
            "os": "linux",
            "kernel": "Linux test",
            "cpu_model": "test CPU",
            "total_memory_bytes": 1_000_000_000,
            "page_size_bytes": 4096,
            "filesystem_type": "tmpfs",
            "source_destination_same_device": True,
            "cpu_affinity": "2",
            "storage_identifier": None,
        },
        "configuration": configuration,
        "parallel_metrics": {
            "schema_version": 1,
            "scope": "explicit_local_execution_only",
            "claim": "descriptive",
            "configured_worker_budget": {
                "status": "measured",
                "value": [1],
                "scope": "configuration.execution_workers",
            },
            "observed_process_thread_count": {
                "status": "unavailable",
                "scope": "process_thread_count",
                "reason": "no process-global thread counter is collected",
            },
            "cases": parallel_cases,
        },
        "results": results,
    }


def _reports():
    return {
        "a1": _report("a" * 40, "a" * 64, "a1"),
        "b1": _report("b" * 40, "b" * 64, "b1"),
        "b2": _report("b" * 40, "b" * 64, "b2"),
        "a2": _report("a" * 40, "a" * 64, "a2"),
    }


class OverlayAbbaSummaryTests(unittest.TestCase):
    def test_validates_matrix_and_summarizes_publication_only(self):
        result = summary.summarize_reports(_reports())

        self.assertEqual(result["verification"]["result_count"], 9)
        self.assertEqual(
            result["protocol"]["timed_metric"],
            "source.opc_source_overlay.publication_ns",
        )
        self.assertEqual(result["protocol"]["samples_per_leg"], summary.SAMPLE_COUNT)
        self.assertEqual(
            result["protocol"]["warmup_iterations_per_leg"],
            summary.WARMUP_ITERATIONS,
        )
        self.assertFalse(result["protocol"]["filesystem_fresh_child_semantics_claimed"])
        self.assertEqual(
            {(item["shape"], item["overlay_count"]) for item in result["results"]},
            {(shape, count) for shape in summary.SHAPES for count in summary.COUNTS},
        )
        self.assertEqual(
            result["results"][0]["publication_ns"]["accepted_statistics"],
            list(summary.STATISTICS),
        )
        self.assertNotIn("elapsed_ns", result["results"][0])
        self.assertFalse(result["verification"]["total_elapsed_statistics_claimed"])
        self.assertFalse(result["verification"]["allocator_timing_claimed"])
        self.assertTrue(result["verification"]["source_counter_vectors_verified"])
        self.assertTrue(result["verification"]["operation_metrics_verified"])

    def test_rejects_phase_sum_mutation(self):
        reports = _reports()
        row = reports["a1"]["results"][0]
        row["source"]["opc_source_overlay"]["publication_ns"][0] += 1
        with self.assertRaises(summary.OverlayAbbaInputError):
            summary.summarize_reports(reports)

    def test_rejects_wrong_matrix_cardinality_and_duplicate_cell(self):
        reports = _reports()
        reports["a1"]["results"].pop()
        with self.assertRaises(summary.OverlayAbbaInputError):
            summary.summarize_reports(reports)

        reports = _reports()
        reports["b1"]["results"][-1] = copy.deepcopy(reports["b1"]["results"][0])
        with self.assertRaises(summary.OverlayAbbaInputError):
            summary.summarize_reports(reports)

        reports = _reports()
        reports["a1"]["results"][0], reports["a1"]["results"][1] = (
            reports["a1"]["results"][1],
            reports["a1"]["results"][0],
        )
        with self.assertRaises(summary.OverlayAbbaInputError):
            summary.summarize_reports(reports)

    def test_rejects_source_sink_oracle_identity_mutations(self):
        for field in ("source_sha256", "expected_eager_sha256"):
            reports = _reports()
            reports["b1"]["results"][0]["source"]["opc_source_overlay"][field] = "c" * 64
            with self.subTest(field=field), self.assertRaises(summary.OverlayAbbaInputError):
                summary.summarize_reports(reports)

        reports = _reports()
        reports["b2"]["results"][0]["sink"]["accepted_bytes"] = 1235
        with self.assertRaises(summary.OverlayAbbaInputError):
            summary.summarize_reports(reports)

        reports = _reports()
        reports["a2"]["results"][0]["source"]["opc_source_overlay"][
            "raw_members_and_order_preservation_verified"
        ] = False
        with self.assertRaises(summary.OverlayAbbaInputError):
            summary.summarize_reports(reports)

    def test_rejects_allocator_and_duplicate_reports(self):
        reports = _reports()
        reports["b1"]["tool"]["instrumentation"] = "system_allocator_operation_scoped"
        with self.assertRaises(summary.OverlayAbbaInputError):
            summary.summarize_reports(reports)

        reports = _reports()
        reports["b2"] = copy.deepcopy(reports["b1"])
        with self.assertRaises(summary.OverlayAbbaInputError):
            summary.summarize_reports(reports)

    def test_rejects_formal_protocol_and_operation_evidence_mutations(self):
        mutations = {
            "samples": lambda report: report["configuration"].update(
                samples_per_case=499
            ),
            "warmups": lambda report: report["configuration"].update(
                warmup_iterations_per_case=19
            ),
            "cache_state": lambda report: report["configuration"].update(
                filesystem_cache_states=["cold"]
            ),
            "workers": lambda report: report["configuration"].update(
                execution_workers=[2]
            ),
            "filesystem_root": lambda report: report["configuration"].update(
                filesystem_root_selected=True
            ),
            "toolchain": lambda report: report["environment"].update(
                rustc_version="rustc 1.97.0 (wrong)"
            ),
            "cpu_affinity": lambda report: report["environment"].update(
                cpu_affinity="0"
            ),
            "timing_scope": lambda report: report["results"][0]["source"][
                "opc_source_overlay"
            ].update(timing_scope="mentions publication_ns but omits the harness contract"),
            "source_vector": lambda report: report["results"][0]["source"].update(
                read_bytes=[value + 1 for value in report["results"][0]["source"]["read_bytes"]]
            ),
            "sample_order": lambda report: report["results"][0]["elapsed_ns"].update(
                sample_order=list(reversed(report["results"][0]["elapsed_ns"]["sample_order"])
                )
            ),
            "operation_null": lambda report: report["results"][0].update(
                operation_metrics=None
            ),
            "operation_missing": lambda report: report["results"][0].pop(
                "operation_metrics"
            ),
            "operation_sample_count": lambda report: report["results"][0][
                "operation_metrics"
            ].update(sample_count=summary.SAMPLE_COUNT - 1),
            "operation_sink": lambda report: report["results"][0][
                "operation_metrics"
            ]["sink"]["accepted_bytes"]["values"].__setitem__(0, 0),
            "allocation_values": lambda report: report["results"][0][
                "operation_metrics"
            ]["allocation"]["allocation_calls"].update(
                values=[0] * summary.SAMPLE_COUNT
            ),
            "allocation_null_values": lambda report: report["results"][0][
                "operation_metrics"
            ]["allocation"]["allocation_calls"].update(values=None),
            "cache_vector": lambda report: report["results"][0]["source"][
                "opc_source_overlay"
            ]["cache_before_publication_hits"].__setitem__(0, 1),
            "sink": lambda report: report["results"][0]["sink"].update(
                largest_write=report["results"][0]["sink"]["largest_write"] + 1
            ),
        }
        for name, mutate in mutations.items():
            reports = _reports()
            with self.subTest(name=name):
                mutate(reports["a1"])
                with self.assertRaises(summary.OverlayAbbaInputError):
                    summary.summarize_reports(reports)

    def test_authorizes_only_positive_pairs_and_drift_within_ceiling(self):
        reports = _reports()
        row = reports["b2"]["results"][0]
        publication = row["source"]["opc_source_overlay"]["publication_ns"]
        row["source"]["opc_source_overlay"]["publication_ns"] = [value + 200 for value in publication]
        elapsed = row["elapsed_ns"]
        elapsed["samples"] = [value + 200 for value in elapsed["samples"]]
        refreshed = _statistics(elapsed["samples"], elapsed["sample_order"])
        elapsed.update(refreshed)
        result = summary.summarize_reports(reports)
        cell = next(item for item in result["results"] if item["shape"] == "overlay-small" and item["overlay_count"] == 2)
        self.assertFalse(cell["publication_ns"]["authorized_statistics"]["p50"])
        self.assertIn("p50", cell["publication_ns"]["rejected_statistics"])

    def test_publication_statistic_edge_math_is_u64_safe(self):
        self.assertEqual(
            summary._statistic([1]),
            {"p50": 1, "mean": 1.0, "p95": 1, "p99": 1},
        )
        self.assertEqual(
            summary._statistic([1, 2]),
            {"p50": 1, "mean": 1.5, "p95": 2, "p99": 2},
        )
        edge = summary._statistic([summary.U64_MAX - 1, summary.U64_MAX])
        self.assertEqual(edge["p50"], summary.U64_MAX - 1)
        self.assertEqual(edge["p95"], summary.U64_MAX)
        self.assertEqual(edge["p99"], summary.U64_MAX)
        self.assertTrue(math.isfinite(edge["mean"]))
        with self.assertRaises(summary.OverlayAbbaInputError):
            summary._statistic([])
        with self.assertRaises(summary.OverlayAbbaInputError):
            summary._statistic([0, 1])

    def test_rejects_unknown_keys_and_strict_json_inputs(self):
        reports = _reports()
        reports["a1"]["results"][0]["unexpected"] = True
        with self.assertRaises(summary.OverlayAbbaInputError):
            summary.summarize_reports(reports)

        reports = _reports()
        reports["extra"] = copy.deepcopy(reports["a1"])
        with self.assertRaises(summary.OverlayAbbaInputError):
            summary.summarize_reports(reports)

        with tempfile.TemporaryDirectory() as directory:
            duplicate = Path(directory) / "duplicate.json"
            duplicate.write_text('{"schema_version":1,"schema_version":1}', encoding="utf-8")
            with self.assertRaises(summary.OverlayAbbaInputError):
                summary.load_report(duplicate)

            nonfinite = Path(directory) / "nonfinite.json"
            nonfinite.write_text('{"value":NaN}', encoding="utf-8")
            with self.assertRaises(summary.OverlayAbbaInputError):
                summary.load_report(nonfinite)

    def test_cli_emits_json_and_rejects_invalid_input(self):
        reports = _reports()
        with tempfile.TemporaryDirectory() as directory:
            paths = []
            for role in summary.ROLES:
                path = Path(directory) / f"{role}.json"
                path.write_text(json.dumps(reports[role]), encoding="utf-8")
                paths.append(path)
            output = Path(directory) / "summary.json"
            self.assertEqual(
                summary.main(
                    [
                        "--a1",
                        str(paths[0]),
                        "--b1",
                        str(paths[1]),
                        "--b2",
                        str(paths[2]),
                        "--a2",
                        str(paths[3]),
                        "--json-out",
                        str(output),
                    ]
                ),
                0,
            )
            encoded = json.loads(output.read_text(encoding="utf-8"))
            self.assertEqual(encoded["verification"]["result_count"], 9)
            self.assertEqual(summary.main([str(path) for path in paths[:3]]), 2)


if __name__ == "__main__":
    unittest.main()
