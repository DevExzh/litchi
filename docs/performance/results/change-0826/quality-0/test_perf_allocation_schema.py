"""Regression tests for the single-report allocation schema admission helper.

The retained 0825 report is used only as a pinned positive compatibility
fixture.  Every malformed and invariant case below is a small in-memory report
with the same shape emitted by the ordinary-save harness, so these tests never
start a workload or reuse a timing sample as a new measurement.
"""

from __future__ import annotations

import hashlib
import io
import json
import tempfile
import unittest
from contextlib import redirect_stderr, redirect_stdout
from pathlib import Path

from tools import perf_allocation_schema


ROOT = Path(__file__).resolve().parents[1]
QUALIFICATION_REPORT = (
    ROOT
    / "docs"
    / "performance"
    / "results"
    / "change-0825"
    / "qualification-before"
    / "00-docx-lifecycle.json"
)
QUALIFICATION_REPORT_SHA256 = (
    "9f7ac2c8402d1565235746fef5eb1076c0bc159fb0908d50954af203908d2ada"
)

CASE = "docx_real_file_ordinary_save_lifecycle"
SAMPLES = 3
WARMUP = 0
SCOPE = perf_allocation_schema.SCOPE
REVISION = perf_allocation_schema.REVISION
U64_MAX = perf_allocation_schema.U64_MAX

EXPECTED_ALLOCATION_FIELDS = (
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


def _metric(values: list[int] | None, *, status: str = "measured") -> dict:
    metric = {"status": status, "scope": SCOPE}
    if values is not None:
        metric["values"] = list(values)
    return metric


def _allocation(*, measured: bool) -> dict:
    if not measured:
        return {
            "status": "unavailable",
            "scope": SCOPE,
            **{
                field: _metric(None, status="unavailable")
                for field in EXPECTED_ALLOCATION_FIELDS
            },
        }

    # Sample 1 deliberately has negative signed net-live bytes:
    # 250 - 300 == 50 - 100 == -50.  The region peak is an operation-region
    # peak and can be below the process-lifetime peak after the operation.
    vectors = {
        "allocation_calls": [10, 20, 0],
        "deallocation_calls": [8, 40, 0],
        "reallocation_calls": [1, 2, 0],
        "failed_allocation_calls": [0, 0, 0],
        "allocated_bytes": [100, 50, 0],
        "deallocated_bytes": [50, 100, 0],
        "live_bytes_before": [200, 300, 0],
        "live_bytes_after": [250, 250, 0],
        "peak_live_bytes_before": [300, 350, 0],
        "peak_live_bytes_after": [400, 450, 0],
        "region_peak_live_bytes": [275, 325, 0],
    }
    return {
        "status": "measured",
        "scope": SCOPE,
        **{field: _metric(vectors[field]) for field in EXPECTED_ALLOCATION_FIELDS},
    }


def _report(
    *,
    mode: str = "observer",
    instrumentation: str | None = None,
    revision: str | None = REVISION,
) -> dict:
    observer = mode == "observer"
    if instrumentation is None:
        instrumentation = (
            "system_allocator_operation_scoped" if observer else "none"
        )
    tool = {
        "name": "litchi-perf-baseline",
        "version": "0.1.0",
        "binary": "litchi-perf-baseline-alloc"
        if observer
        else "litchi-perf-baseline",
        "profile": "release",
        "target_os": "linux",
        "target_arch": "x86_64",
        "instrumentation": instrumentation,
    }
    if revision is not None:
        tool["allocator_counter_revision"] = revision

    elapsed_samples = [100, 100, 200]
    operation_metrics = {
        "sample_count": SAMPLES,
        "sample_indices": [0, 1, 2],
        "alignment": perf_allocation_schema.ALIGNMENT,
        "latency_claim": (
            "allocator_instrumented_elapsed_not_latency_claim"
            if observer
            else "comparable_timed_operation"
        ),
        # These fields mirror the emitted operation_metrics shape.  The
        # allocation schema helper intentionally leaves their own contracts to
        # the corresponding readers.
        "source": {"status": "unavailable", "scope": "source"},
        "process": {"status": "unavailable", "scope": "process"},
        "sink": {"status": "unavailable", "scope": "sink"},
        "publication": {"status": "unavailable", "scope": "publication"},
        "materialization": {
            "status": "unavailable",
            "scope": "materialization",
        },
        "cfb_phases": {"status": "unavailable", "scope": "cfb_phases"},
        "allocation": _allocation(measured=observer),
    }
    return {
        "schema_version": 1,
        "tool": tool,
        "configuration": {
            "samples_per_case": SAMPLES,
            "warmup_iterations_per_case": WARMUP,
            "cases": [CASE],
        },
        "results": [
            {
                "case": CASE,
                "elapsed_ns": {
                    "unit": "ns",
                    "samples": elapsed_samples,
                    "sample_order": [0, 1, 2],
                },
                "operation_metrics": operation_metrics,
            }
        ],
    }


def _validate(report: dict, *, mode: str = "observer") -> None:
    perf_allocation_schema.validate_report(
        report,
        expected_case=CASE,
        expected_samples=SAMPLES,
        expected_warmup=WARMUP,
        mode=mode,
    )


class RetainedQualificationAcceptanceTests(unittest.TestCase):
    def test_retained_0825_qualification_report_is_accepted(self) -> None:
        # The digest is the frozen identity.  Parse the report only after the
        # bytes have been checked, so a changed retained artifact cannot silently
        # become the compatibility fixture.
        raw = QUALIFICATION_REPORT.read_bytes()
        self.assertEqual(
            hashlib.sha256(raw).hexdigest(), QUALIFICATION_REPORT_SHA256
        )
        report = json.loads(raw.decode("utf-8"))
        perf_allocation_schema.validate_report(
            report,
            expected_case=CASE,
            expected_samples=1,
            expected_warmup=0,
            mode="observer",
        )


class PositiveFixtureTests(unittest.TestCase):
    def test_expected_field_set_is_the_pinned_eleven_field_v3_vector(self) -> None:
        self.assertEqual(
            tuple(perf_allocation_schema.ALLOCATION_FIELDS),
            EXPECTED_ALLOCATION_FIELDS,
        )

    def test_observer_accepts_signed_negative_net_live_and_distinct_peaks(self) -> None:
        report = _report()
        allocation = report["results"][0]["operation_metrics"]["allocation"]
        before = allocation["live_bytes_before"]["values"][1]
        after = allocation["live_bytes_after"]["values"][1]
        allocated = allocation["allocated_bytes"]["values"][1]
        deallocated = allocation["deallocated_bytes"]["values"][1]
        self.assertEqual(after - before, -50)
        self.assertEqual(after - before, allocated - deallocated)
        self.assertLess(
            allocation["region_peak_live_bytes"]["values"][0],
            allocation["peak_live_bytes_after"]["values"][0],
        )
        _validate(report)

    def test_nonzero_failed_allocation_calls_are_valid_schema_data(self) -> None:
        report = _report()
        report["results"][0]["operation_metrics"]["allocation"][
            "failed_allocation_calls"
        ]["values"] = [2, 1, 0]
        _validate(report)


class AllocationVectorMalformedTests(unittest.TestCase):
    def _assert_rejected(self, mutate) -> None:
        report = _report()
        allocation = report["results"][0]["operation_metrics"]["allocation"]
        mutate(allocation)
        with self.assertRaises(perf_allocation_schema.AllocationSchemaError):
            _validate(report)

    def test_every_v3_field_rejects_each_malformed_vector_shape(self) -> None:
        def malformed(kind: str, field: str):
            def mutate(allocation: dict) -> None:
                if kind == "scalar":
                    allocation[field] = 7
                elif kind == "short":
                    allocation[field]["values"] = allocation[field]["values"][:-1]
                elif kind == "long":
                    allocation[field]["values"] = allocation[field]["values"] + [0]
                elif kind == "null":
                    allocation[field]["values"] = None
                elif kind == "bool":
                    allocation[field]["values"] = [True] * SAMPLES
                elif kind == "float":
                    allocation[field]["values"] = [1.0] * SAMPLES
                elif kind == "negative":
                    allocation[field]["values"] = [-1] * SAMPLES
                elif kind == "u64overflow":
                    allocation[field]["values"] = [U64_MAX + 1] * SAMPLES
                elif kind == "status":
                    allocation[field]["status"] = "unavailable"
                elif kind == "scope":
                    allocation[field]["scope"] = "unexpected_scope"
                elif kind == "unknown":
                    allocation[field]["unexpected"] = True
                elif kind == "missing":
                    del allocation[field]
                else:  # pragma: no cover - keeps a typo in this table loud.
                    raise AssertionError(kind)

            return mutate

        kinds = (
            "scalar",
            "short",
            "long",
            "null",
            "bool",
            "float",
            "negative",
            "u64overflow",
            "status",
            "scope",
            "unknown",
            "missing",
        )
        for field in EXPECTED_ALLOCATION_FIELDS:
            for kind in kinds:
                with self.subTest(field=field, malformed=kind):
                    self._assert_rejected(malformed(kind, field))

    def test_missing_values_key_is_rejected_as_schema_error(self) -> None:
        def mutate(allocation: dict) -> None:
            allocation["allocated_bytes"].pop("values")

        self._assert_rejected(mutate)

    def test_allocation_byte_conservation_is_required(self) -> None:
        def mutate(allocation: dict) -> None:
            allocation["allocated_bytes"]["values"][0] += 1

        self._assert_rejected(mutate)

    def test_reallocation_calls_cannot_exceed_allocation_calls(self) -> None:
        def mutate(allocation: dict) -> None:
            allocation["reallocation_calls"]["values"][0] = 11

        self._assert_rejected(mutate)

    def test_peak_before_must_cover_live_before(self) -> None:
        def mutate(allocation: dict) -> None:
            allocation["peak_live_bytes_before"]["values"][0] = 199

        self._assert_rejected(mutate)

    def test_peak_after_cannot_decrease_from_peak_before(self) -> None:
        def mutate(allocation: dict) -> None:
            allocation["peak_live_bytes_after"]["values"][0] = 299

        self._assert_rejected(mutate)

    def test_region_peak_must_cover_both_live_endpoints(self) -> None:
        def mutate(allocation: dict) -> None:
            allocation["region_peak_live_bytes"]["values"][0] = 199

        self._assert_rejected(mutate)

    def test_region_peak_cannot_exceed_lifetime_peak_after(self) -> None:
        def mutate(allocation: dict) -> None:
            allocation["region_peak_live_bytes"]["values"][0] = 401

        self._assert_rejected(mutate)


class NativeUnavailableTests(unittest.TestCase):
    def test_native_report_accepts_unavailable_vectors_without_numeric_payload(self) -> None:
        report = _report(mode="native", revision=None)
        _validate(report, mode="native")
        allocation = report["results"][0]["operation_metrics"]["allocation"]
        for field in EXPECTED_ALLOCATION_FIELDS:
            self.assertNotIn("values", allocation[field])

    def test_native_report_rejects_a_numeric_payload_for_every_vector(self) -> None:
        for field in EXPECTED_ALLOCATION_FIELDS:
            with self.subTest(field=field):
                report = _report(mode="native", revision=None)
                report["results"][0]["operation_metrics"]["allocation"][field][
                    "values"
                ] = [0] * SAMPLES
                with self.assertRaises(perf_allocation_schema.AllocationSchemaError):
                    _validate(report, mode="native")

    def test_native_tool_must_omit_allocator_counter_revision(self) -> None:
        report = _report(mode="native")
        with self.assertRaises(perf_allocation_schema.AllocationSchemaError):
            _validate(report, mode="native")


class FrozenIdentityTests(unittest.TestCase):
    def test_case_sample_and_warmup_expectations_are_supplied_contracts(self) -> None:
        for keyword, value in (
            ("expected_case", "another_case"),
            ("expected_samples", SAMPLES - 1),
            ("expected_warmup", 1),
        ):
            with self.subTest(expectation=keyword):
                report = _report()
                kwargs = {
                    "expected_case": CASE,
                    "expected_samples": SAMPLES,
                    "expected_warmup": WARMUP,
                    "mode": "observer",
                }
                kwargs[keyword] = value
                with self.assertRaises(perf_allocation_schema.AllocationSchemaError):
                    perf_allocation_schema.validate_report(report, **kwargs)

    def test_only_native_and_observer_are_valid_modes(self) -> None:
        for mode in ("", "allocator", None):
            with self.subTest(mode=mode):
                with self.assertRaises(perf_allocation_schema.AllocationSchemaError):
                    perf_allocation_schema.validate_report(
                        _report(),
                        expected_case=CASE,
                        expected_samples=SAMPLES,
                        expected_warmup=WARMUP,
                        mode=mode,
                    )

    def test_mode_binds_binary_and_latency_claim(self) -> None:
        native = _report(mode="native", revision=None)
        _validate(native, mode="native")
        with self.assertRaises(perf_allocation_schema.AllocationSchemaError):
            _validate(native, mode="observer")
        observer = _report()
        _validate(observer)
        with self.assertRaises(perf_allocation_schema.AllocationSchemaError):
            _validate(observer, mode="native")


class InstrumentationRevisionTests(unittest.TestCase):
    def test_both_pinned_observer_instrumentation_names_are_accepted(self) -> None:
        for instrumentation in (
            "system_allocator_operation_scoped",
            "ordinary_save_procfs_and_system_allocator_operation_scoped",
        ):
            with self.subTest(instrumentation=instrumentation):
                _validate(_report(instrumentation=instrumentation))

    def test_observer_requires_the_pinned_counter_revision(self) -> None:
        for revision in (None, "serialized_region_peak_v2", 3):
            with self.subTest(revision=revision):
                with self.assertRaises(perf_allocation_schema.AllocationSchemaError):
                    _validate(_report(revision=revision))

    def test_observer_rejects_unknown_instrumentation(self) -> None:
        with self.assertRaises(perf_allocation_schema.AllocationSchemaError):
            _validate(_report(instrumentation="system_allocator_global"))

    def test_native_requires_none_instrumentation(self) -> None:
        for instrumentation in (
            "system_allocator_operation_scoped",
            "ordinary_save_procfs_and_system_allocator_operation_scoped",
        ):
            with self.subTest(instrumentation=instrumentation):
                with self.assertRaises(perf_allocation_schema.AllocationSchemaError):
                    _validate(
                        _report(
                            mode="native",
                            instrumentation=instrumentation,
                            revision=None,
                        ),
                        mode="native",
                    )


class SampleAlignmentTests(unittest.TestCase):
    def _assert_rejected(self, mutate) -> None:
        report = _report()
        mutate(report["results"][0])
        with self.assertRaises(perf_allocation_schema.AllocationSchemaError):
            _validate(report)

    def test_sample_indices_must_equal_sample_order(self) -> None:
        def mutate(result: dict) -> None:
            result["operation_metrics"]["sample_indices"] = [0, 2, 1]

        self._assert_rejected(mutate)

    def test_boolean_sample_order_index_is_not_an_integer(self) -> None:
        def mutate(result: dict) -> None:
            result["elapsed_ns"]["sample_order"] = [True, 1, 2]

        self._assert_rejected(mutate)

    def test_boolean_operation_sample_index_is_not_an_integer(self) -> None:
        def mutate(result: dict) -> None:
            result["operation_metrics"]["sample_indices"] = [True, 1, 2]

        self._assert_rejected(mutate)

    def test_duplicate_sample_order_is_not_a_permutation(self) -> None:
        def mutate(result: dict) -> None:
            result["elapsed_ns"]["sample_order"] = [0, 0, 2]

        self._assert_rejected(mutate)

    def test_tied_elapsed_samples_must_follow_original_index_order(self) -> None:
        def mutate(result: dict) -> None:
            result["elapsed_ns"]["sample_order"] = [1, 0, 2]

        self._assert_rejected(mutate)

    def test_elapsed_samples_must_be_sorted_before_tie_check(self) -> None:
        def mutate(result: dict) -> None:
            result["elapsed_ns"]["samples"] = [200, 100, 100]

        self._assert_rejected(mutate)


class CliTests(unittest.TestCase):
    def _run_cli(self, report: dict) -> tuple[int, str, str]:
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "report.json"
            path.write_text(json.dumps(report), encoding="utf-8")
            stdout = io.StringIO()
            stderr = io.StringIO()
            with redirect_stdout(stdout), redirect_stderr(stderr):
                status = perf_allocation_schema.main(
                    [
                        str(path),
                        "--case",
                        CASE,
                        "--samples",
                        str(SAMPLES),
                        "--warmup",
                        str(WARMUP),
                        "--mode",
                        "observer",
                    ]
                )
            return status, stdout.getvalue(), stderr.getvalue()

    def test_cli_accepts_valid_report(self) -> None:
        status, stdout, stderr = self._run_cli(_report())
        self.assertEqual(status, 0)
        self.assertIn("allocation schema PASS", stdout)
        self.assertEqual(stderr, "")

    def test_cli_reports_schema_failure_and_nonzero_status(self) -> None:
        report = _report()
        report["results"][0]["operation_metrics"]["allocation"][
            "allocated_bytes"
        ]["values"][0] += 1
        status, stdout, stderr = self._run_cli(report)
        self.assertEqual(status, 1)
        self.assertEqual(stdout, "")
        self.assertIn("allocation schema rejected", stderr)


if __name__ == "__main__":
    unittest.main()
