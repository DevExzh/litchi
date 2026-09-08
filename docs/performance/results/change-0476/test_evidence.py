#!/usr/bin/env python3
"""Small mutation tests for the standalone 0476 evidence boundary."""

from __future__ import annotations

import copy
import hashlib
import json
from pathlib import Path
import tempfile
import unittest

import analyze
import report_checks
import verify


def elapsed() -> dict[str, object]:
    return {
        "unit": "ns",
        "samples": [100] * report_checks.SAMPLES,
        "sample_order": list(range(report_checks.SAMPLES)),
        "min": 100,
        "p50": 100,
        "p95": 100,
        "p99": 100,
        "max": 100,
        "mean": 100.0,
        "standard_deviation": 0.0,
        "confidence_interval_95": {
            "method": "two-sided Student's t interval for the mean",
            "lower": 100.0,
            "upper": 100.0,
        },
    }


def metric(values: list[int]) -> dict[str, object]:
    return {"status": "measured", "scope": "operation_global_system_allocator", "values": values}


def report(shape: str, *, allocator: bool = False, output: str = "a" * 64) -> dict[str, object]:
    slides = report_checks.SHAPES[shape]
    row: dict[str, object] = {
        "case": report_checks.CASE,
        "corpus": {
            "shape": shape,
            "slide_count": slides,
            "archive_sha256": output,
        },
        "output_sha256": output,
        "elapsed_ns": elapsed(),
        "sink": {"accepted_bytes": slides + 1024, "write_calls": 9, "largest_write": 512},
        "source": {"pptx_slides": {"slide_count": slides, "text_box_count": slides}},
        "configuration": {
            "samples_per_case": report_checks.SAMPLES,
            "warmup_iterations_per_case": report_checks.WARMUPS,
            "execution_workers": [report_checks.WORKERS],
        },
    }
    if allocator:
        operation = {"sample_count": report_checks.SAMPLES, "sample_indices": list(range(report_checks.SAMPLES))}
        operation["allocation"] = {
            "status": "measured",
            "scope": "operation_global_system_allocator",
            "allocation_calls": metric([3] * report_checks.SAMPLES),
            "deallocation_calls": metric([1] * report_checks.SAMPLES),
            "reallocation_calls": metric([1] * report_checks.SAMPLES),
            "failed_allocation_calls": metric([0] * report_checks.SAMPLES),
            "allocated_bytes": metric([50] * report_checks.SAMPLES),
            "deallocated_bytes": metric([50] * report_checks.SAMPLES),
            "live_bytes_before": metric([100] * report_checks.SAMPLES),
            "live_bytes_after": metric([100] * report_checks.SAMPLES),
            "peak_live_bytes_before": metric([100] * report_checks.SAMPLES),
            "peak_live_bytes_after": metric([125] * report_checks.SAMPLES),
            "region_peak_live_bytes": metric([125] * report_checks.SAMPLES),
        }
        row["operation_metrics"] = operation
    return {
        "results": [row],
        "configuration": {
            "samples_per_case": report_checks.SAMPLES,
            "warmup_iterations_per_case": report_checks.WARMUPS,
            "execution_workers": [report_checks.WORKERS],
        },
    }


def formal_protocol() -> dict[str, object]:
    rows: list[dict[str, object]] = []
    for repeat, cases, arms in (
        ("R1", [(mode, shape) for mode in analyze.MODES for shape in analyze.SHAPES], analyze.ARMS),
        ("R2", list(reversed([(mode, shape) for mode in analyze.MODES for shape in analyze.SHAPES])), tuple(reversed(analyze.ARMS))),
    ):
        for mode, shape in cases:
            for arm in arms:
                rows.append({"lane": f"{repeat}-{arm}-{mode}-{shape}", "kind": "formal", "repeat": repeat, "arm": arm, "mode": mode, "shape": shape, "samples": 30, "warmups": 3})
    return {
        "schema": analyze.PROTOCOL_SCHEMA,
        "selector": report_checks.CASE,
        "samples": 30,
        "warmups": 3,
        "workers": 1,
        "cpu": 2,
        "pilot_samples": 1,
        "pilot_warmups": 0,
        "large_requested_bytes_reduction_required_percent": 95,
        "normal_regression_review_threshold_percent": 5,
        "normal_repeat_review_threshold_percent": 5,
        "peak_rss_regression_review_threshold_percent": 5,
        "order": rows,
        "acceptance": {"requested_bytes_reduction_percent": 95},
    }


class EvidenceTests(unittest.TestCase):
    def test_elapsed_mutation_is_rejected(self) -> None:
        value = elapsed()
        value["mean"] = 99
        with self.assertRaises(report_checks.ReportError):
            report_checks.check_elapsed(value)

    def test_allocator_balance_mutation_is_rejected(self) -> None:
        value = report("tiny", allocator=True)["results"][0]["operation_metrics"]["allocation"]
        value["allocated_bytes"]["values"][0] = 0
        with self.assertRaises(report_checks.ReportError):
            report_checks.check_allocation(value, [100] * report_checks.SAMPLES, "mutated.allocation")

    def test_normal_report_cannot_claim_measured_allocator_vectors(self) -> None:
        lane = {"mode": "normal", "shape": "tiny"}
        value = report("tiny", allocator=True)
        with self.assertRaises(report_checks.ReportError):
            report_checks.validate_report(value, lane)

    def test_protocol_requires_reversed_abba_formal_order(self) -> None:
        protocol = formal_protocol()
        analyze.protocol_order(protocol)
        protocol["order"][-1]["arm"] = "candidate"
        with self.assertRaises(analyze.AnalysisError):
            analyze.protocol_order(protocol)

    def test_counter_parser_retains_perf_scaling_metadata(self) -> None:
        text = "123;;cycles:u;4567;83.00;;\n<not supported>;;LLC-load-misses:u;0;100.00;;\n"
        parsed = report_checks.parse_counter_text(text, ["cycles:u", "LLC-load-misses:u"])
        self.assertEqual(parsed["cycles:u"]["event_runtime_ns"], 4567)
        self.assertEqual(parsed["cycles:u"]["running_percent"], 83.0)
        self.assertEqual(parsed["LLC-load-misses:u"]["status"], "not_supported")

    def test_duplicate_json_keys_are_rejected(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "duplicate.json"
            path.write_text('{"a":1,"a":2}\n', encoding="utf-8")
            with self.assertRaises(analyze.AnalysisError):
                analyze.read_json(path)

    def test_source_manifest_rejects_traversal_and_bad_hash(self) -> None:
        with self.assertRaises(verify.VerificationError):
            verify.normalize_manifest({"../outside.rs": "0" * 64}, "manifest")
        with self.assertRaises(verify.VerificationError):
            verify.normalize_manifest({"src/lib.rs": "bad"}, "manifest")

    def test_receipt_artifact_mutation_is_rejected(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            payload = root / "stdout.log"
            payload.write_text("captured\n", encoding="utf-8")
            metadata = {"sha256": hashlib.sha256(payload.read_bytes()).hexdigest(), "bytes": payload.stat().st_size}
            verify.verify_artifacts(root, {"stdout.log": metadata}, "lane.receipt")
            payload.write_text("mutated\n", encoding="utf-8")
            with self.assertRaises(verify.VerificationError):
                verify.verify_artifacts(root, {"stdout.log": metadata}, "lane.receipt")

    def test_rust_validation_ledger_binds_required_and_retained_receipts(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            validation = root / "validation"
            validation.mkdir()
            (root / "candidate-source.json").write_text('{"arm":"candidate"}\n', encoding="utf-8")
            (root / "format-exception.json").write_text(
                json.dumps(
                    {
                        "identical_to_pre_candidate_head": True,
                        "path": "crates/litchi-keynote/src/document.rs",
                        "pre_candidate_head": "a" * 40,
                        "sha256": "b" * 64,
                    }
                ),
                encoding="utf-8",
            )
            source = {"files": 1, "path": "validation-sources/source.json", "sha256": "c" * 64}
            records = {
                "required": {"exit_code": 0, "source_before": source},
                "retained": {"exit_code": 101, "source_before": source},
            }
            ledger: dict[str, object] = {}
            for label, record in records.items():
                receipt = validation / f"{label}.json"
                receipt.write_text(json.dumps(record), encoding="utf-8")
                ledger[label] = {
                    "exit_code": record["exit_code"],
                    "receipt_sha256": hashlib.sha256(receipt.read_bytes()).hexdigest(),
                    "source_before": source,
                }
            (root / "rust-validation.json").write_text(
                json.dumps(
                    {
                        "schema": "litchi-0476-rust-validation-v1",
                        "candidate_source": "candidate-source.json",
                        "format_exception": "format-exception.json",
                        "required_success": ["required"],
                        "retained_nonzero_attempts": {"retained": 101},
                        "receipts": ledger,
                    }
                ),
                encoding="utf-8",
            )
            self.assertEqual(verify.verify_rust_validation(root)["receipts"], 2)
            (validation / "required.json").write_text('{"exit_code":0,"source_before":null}', encoding="utf-8")
            with self.assertRaises(verify.VerificationError):
                verify.verify_rust_validation(root)

    def test_allocator_operation_peak_uses_baseline_adjusted_delta(self) -> None:
        def checked(region: list[int], before: list[int]) -> dict[str, object]:
            allocation = {
                field: [0, 0]
                for field in report_checks.ALLOCATOR_FIELDS
            }
            allocation["region_peak_live_bytes"] = region
            allocation["live_bytes_before"] = before
            allocation["live_bytes_after"] = before[:]
            return {
                "lane": "test",
                "elapsed": [100, 100],
                "elapsed_summary": {
                    "min": 100,
                    "p50": 100,
                    "p95": 100,
                    "p99": 100,
                    "max": 100,
                    "mean": 100.0,
                    "standard_deviation": 0.0,
                    "confidence_interval_95": {"lower": 100.0, "upper": 100.0},
                },
                "peak_rss_kib": 10,
                "operation": {"allocation": allocation},
            }

        control = checked([1_010, 1_010], [1_000, 1_000])
        candidate = checked([2_005, 2_005], [2_000, 2_000])
        pair = analyze._pair_summary(control, candidate, "allocator", "tiny", "R1")
        self.assertEqual(pair["allocation"]["region_peak_live_bytes"]["control"]["mean"], 1_010)
        self.assertEqual(pair["allocation"]["peak_live_bytes_delta"]["control"]["mean"], 10)
        self.assertEqual(pair["allocation"]["peak_live_bytes_delta"]["candidate"]["mean"], 5)
        self.assertFalse(pair["operation_peak_review_required"])
        rows = {}
        for repeat in analyze.REPEATS:
            for arm in analyze.ARMS:
                for mode in analyze.MODES:
                    for shape in analyze.SHAPES:
                        rows[(repeat, arm, mode, shape)] = checked(
                            [1_010, 1_010] if repeat == "R1" else [1_100, 1_100],
                            [1_000, 1_000],
                        )
        drift = analyze._repeat_summary(rows)["control-allocator-tiny"]
        self.assertTrue(drift["operation_peak_exceeds_review_threshold"])

    def test_seal_requires_exact_regular_file_coverage(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            payload = root / "payload.txt"
            payload.write_text("evidence\n", encoding="utf-8")
            digest = hashlib.sha256(payload.read_bytes()).hexdigest()
            (root / "SHA256SUMS").write_text(f"{digest}  payload.txt\n", encoding="utf-8")
            self.assertEqual(verify.verify_seal(root)["status"], "pass")
            (root / "extra.txt").write_text("unsealed\n", encoding="utf-8")
            with self.assertRaises(verify.VerificationError):
                verify.verify_seal(root)


if __name__ == "__main__":
    unittest.main(verbosity=2)
