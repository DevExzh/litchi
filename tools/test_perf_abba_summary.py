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
                "operation_metrics identity",
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

    def test_allocator_instrumentation_is_not_accepted_for_latency_abba(self):
        legs = four_legs()
        for leg in legs:
            leg["tool"]["instrumentation"] = "system_allocator"
        with self.assertRaisesRegex(
            perf_abba_summary.AbbaSummaryInputError,
            "instrumentation.*latency ABBA",
        ):
            perf_abba_summary.summarize_reports(legs)

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


if __name__ == "__main__":
    unittest.main()
