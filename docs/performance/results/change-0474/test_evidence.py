#!/usr/bin/env python3
"""Focused, repository-free tests for the 0474 evidence boundary."""

from __future__ import annotations

import hashlib
import copy
import report_checks
import json
import tempfile
import unittest
from pathlib import Path
import sys


sys.path.insert(0, str(Path(__file__).resolve().parent))
import analyze  # noqa: E402
import verify  # noqa: E402


def row(shape: str, *, output: str, write_calls: int = 9) -> dict:
    slides = analyze.SLIDES[shape]
    return {
        "case": analyze.CASE,
        "corpus": {"shape": shape, "slide_count": slides},
        "elapsed_ns": {"samples": list(range(1, analyze.SAMPLES + 1)),
                        "sample_order": list(range(analyze.SAMPLES))},
        "output_sha256": output,
        "sink": {
            "accepted_bytes": 1024 + slides,
            "write_calls": write_calls,
            "input_bytes": slides * 12,
            "authored_part_bytes": slides * 64,
        },
        "source": {"pptx_slides": {
            "slide_count": slides,
            "text_box_count": slides,
            "input_text_bytes": slides * 12,
            "authored_part_bytes": slides * 64,
            "observed_max_slide_xml_bytes": 128,
            "max_slide_xml_bytes": 16_384,
            "structural_metadata_fixed_member_count": 37,
            "structural_metadata_members_per_slide": 2,
        }},
    }


def allocation_row(shape: str, *, output: str) -> dict:
    result = row(shape, output=output)
    vectors = {field: [index + 1 for index in range(analyze.SAMPLES)]
               for field in analyze.ALLOCATOR_FIELDS}
    vectors["live_bytes_before"] = [100] * analyze.SAMPLES
    vectors["live_bytes_after"] = [102] * analyze.SAMPLES
    vectors["peak_live_bytes_before"] = [300] * analyze.SAMPLES
    vectors["peak_live_bytes_after"] = [330] * analyze.SAMPLES
    vectors["region_peak_live_bytes"] = [330] * analyze.SAMPLES
    result["operation_metrics"] = {
        "allocation": {"status": "measured", "scope": "operation_global_system_allocator", **vectors}
    }
    return result


class EvidenceTests(unittest.TestCase):
    def test_real_report_semantic_and_statistic_mutations_are_rejected(self) -> None:
        lane = {"mode": "normal", "repeat": "R1", "shape": "tiny", "lane": "R1-normal-tiny"}
        report = analyze.read_json(analyze.ROOT / "captures/R1-normal-tiny/report.json")
        analyze.validate_report(report, lane, "original")
        mutations = [
            lambda r: r["results"][0]["elapsed_ns"].update(mean=0),
            lambda r: r["results"][0]["source"]["pptx_slides"].update(semantic_sha256="0" * 64),
            lambda r: r["results"][0]["source"].pop("pptx_slides"),
            lambda r: r["results"][0]["sink"].update(retained_output_bytes=1),
            lambda r: r["results"][0]["source"]["pptx_slides"].update(slide_count=7),
            lambda r: r["results"][0]["source"]["pptx_slides"].update(text_box_count=7),
            lambda r: r["results"][0]["source"]["pptx_slides"].update(structural_metadata_fixed_member_count=38),
            lambda r: r["results"][0]["source"]["pptx_slides"].update(observed_max_slide_xml_bytes=0),
            lambda r: r["configuration"].update(samples_per_case=29),
        ]
        for mutate in mutations:
            changed = copy.deepcopy(report); mutate(changed)
            with self.subTest(mutation=mutate), self.assertRaises(analyze.AnalysisError):
                analyze.validate_report(changed, lane, "mutated")

    def test_real_allocator_byte_balance_failure_and_scope_are_rejected(self) -> None:
        lane = {"mode": "allocator", "repeat": "R1", "shape": "tiny", "lane": "R1-allocator-tiny"}
        report = analyze.read_json(analyze.ROOT / "captures/R1-allocator-tiny/report.json")
        analyze.validate_report(report, lane, "original")
        for field, replacement in [("allocated_bytes", 0), ("failed_allocation_calls", 1), ("region_peak_live_bytes", 0)]:
            changed = copy.deepcopy(report)
            changed["results"][0]["operation_metrics"]["allocation"][field]["values"][0] = replacement
            with self.subTest(field=field), self.assertRaises(analyze.AnalysisError):
                analyze.validate_report(changed, lane, "mutated")
        changed = copy.deepcopy(report)
        changed["results"][0]["operation_metrics"]["allocation"]["scope"] = "modeled_scratch"
        with self.assertRaises(analyze.AnalysisError):
            analyze.validate_report(changed, lane, "mutated scope")

    def test_frozen_protocol_contains_reversed_twelve_lane_matrix(self) -> None:
        protocol = analyze.read_json(analyze.ROOT / "protocol.json")
        order = analyze.protocol_order(protocol)
        self.assertEqual(len(order), 12)
        self.assertEqual(order[0]["lane"], "R1-normal-tiny")
        self.assertEqual(order[-1]["lane"], "R2-normal-tiny")

    def test_signed_allocator_deltas_are_retained(self) -> None:
        checked = allocation_row("tiny", output="a" * 64)
        checked["operation_metrics"]["allocation"]["live_bytes_after"] = [99] * analyze.SAMPLES
        values = analyze.allocation_vectors(checked, "allocator-tiny")
        self.assertEqual(values["live_bytes_delta"], [-1] * analyze.SAMPLES)
        self.assertEqual(analyze.signed_stats(values["live_bytes_delta"], "delta")["p50"], -1)

    def test_allocator_vector_mutation_is_rejected(self) -> None:
        checked = allocation_row("tiny", output="a" * 64)
        checked["operation_metrics"]["allocation"]["allocation_calls"] = [1] * (analyze.SAMPLES - 1)
        with self.assertRaises(analyze.AnalysisError):
            analyze.allocation_vectors(checked, "allocator-tiny")

    def test_identity_matrix_requires_equal_output_sink_and_writer_projection(self) -> None:
        rows = {}
        for mode in analyze.MODES:
            for repeat in analyze.REPEATS:
                for shape in analyze.SHAPES:
                    output = hashlib.sha256(shape.encode()).hexdigest()
                    rows[(mode, repeat, shape)] = row(shape, output=output)
        identities = analyze.validate_identity_matrix(rows)
        self.assertEqual(set(identities), set(analyze.SHAPES))
        rows[("allocator", "R2", "large")] = row("large", output="b" * 64)
        with self.assertRaisesRegex(analyze.AnalysisError, "deterministic writer identity"):
            analyze.validate_identity_matrix(rows)

    def test_identity_matrix_rejects_sink_counter_mutation(self) -> None:
        rows = {}
        for mode in analyze.MODES:
            for repeat in analyze.REPEATS:
                for shape in analyze.SHAPES:
                    output = hashlib.sha256(shape.encode()).hexdigest()
                    rows[(mode, repeat, shape)] = row(shape, output=output)
        rows[("normal", "R2", "medium")] = row(
            "medium", output=hashlib.sha256(b"medium").hexdigest(), write_calls=10
        )
        with self.assertRaises(analyze.AnalysisError):
            analyze.validate_identity_matrix(rows)

    def test_duplicate_json_keys_are_rejected(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "duplicate.json"
            path.write_text('{"a": 1, "a": 2}\n', encoding="utf-8")
            with self.assertRaises(analyze.AnalysisError):
                analyze.read_json(path)

    def test_rss_parser_requires_one_whole_process_value(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "resource.log"
            path.write_text("Maximum resident set size (kbytes): 42\n", encoding="utf-8")
            self.assertEqual(analyze.parse_rss(path, Path(directory))["maximum_resident_set_kib"], 42)
            path.write_text("Maximum resident set size (kbytes): 42\n"
                             "Maximum resident set size (kbytes): 43\n", encoding="utf-8")
            with self.assertRaises(analyze.AnalysisError):
                analyze.parse_rss(path, Path(directory))

    def test_chronology_rejects_capture_before_build(self) -> None:
        build = {"started_utc": "2026-09-08T09:00:00+00:00", "finished_utc": "2026-09-08T09:01:00+00:00"}
        captures = [{"started_utc": "2026-09-08T09:00:59+00:00", "finished_utc": "2026-09-08T09:01:01+00:00"}]
        with self.assertRaises(verify.VerificationError):
            verify.verify_chronology(build, captures)

    def test_seal_requires_exact_regular_file_coverage(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            target = root / "nested.txt"
            target.write_text("evidence\n", encoding="utf-8")
            digest = hashlib.sha256(target.read_bytes()).hexdigest()
            (root / "SHA256SUMS").write_text(f"{digest}  nested.txt\n", encoding="utf-8")
            self.assertEqual(verify.verify_seal(root)["status"], "pass")
            (root / "extra.txt").write_text("unsealed\n", encoding="utf-8")
            with self.assertRaises(verify.VerificationError):
                verify.verify_seal(root)


if __name__ == "__main__":
    unittest.main(verbosity=2)
