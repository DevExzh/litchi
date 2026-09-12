"""Tests for descriptive per-process uncertainty summaries."""

import json
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import patch

import uncertainty


class UncertaintySummaryTests(unittest.TestCase):
    def setUp(self):
        temporary = TemporaryDirectory(prefix="litchi-xlsx-uncertainty-")
        self.addCleanup(temporary.cleanup)
        self.results = Path(temporary.name)

    def write_receipt(
        self,
        process: int,
        sample_count: int = 20,
        lane: str = "capture",
        expected_success: bool = True,
        error_class: str | None = None,
    ):
        samples = [
            {
                "elapsed_ns": index + process,
                "requested_alloc_bytes": 100 + index + process,
                "peak_live_delta": 200 + index + process,
                "expected_success": expected_success,
                "actual_success": expected_success,
                "semantic_ok": True,
                "output_exact": True,
                "alloc_balance_ok": True,
                "alloc_invalid": False,
                "alloc_failed": 0,
                "error": None if expected_success else {"class": error_class},
            }
            for index in range(sample_count)
        ]
        receipt = {
            "schema": "xlsx-svg-lifecycle-profile-v1",
            "lane": lane,
            "warmup": 2,
            "sample_count": sample_count,
            "samples": samples,
            "expected_success": expected_success,
            "input_bytes": 42,
            "input_sha256": "a" * 64,
        }
        (self.results / f"{lane}-p{process}.json").write_text(json.dumps(receipt))
        (self.results / f"{lane}-p{process}.time.txt").write_text(
            f"Maximum resident set size (kbytes): {100 + process * 10}\n"
        )

    def test_medians_ranges_and_between_process_range_are_reproducible(self):
        for process in range(1, 4):
            self.write_receipt(process)
        with patch.object(uncertainty, "LANES", ("capture",)):
            first = uncertainty.summarize_results(self.results)
            second = uncertainty.summarize_results(self.results)
        self.assertEqual(first, second)
        lane = first["lanes"][0]
        self.assertEqual(lane["process_count"], 3)
        self.assertEqual(lane["sample_count_per_process"], 20)
        self.assertEqual(lane["processes"][0]["metrics"]["elapsed_ns"]["median"], 10.5)
        self.assertEqual(lane["processes"][0]["metrics"]["elapsed_ns"]["range"], [1, 20])
        self.assertEqual(
            lane["between_process_medians"]["elapsed_ns"]["range"],
            [10.5, 12.5],
        )
        self.assertEqual(lane["rss_kib"]["process_values"], [110, 120, 130])
        self.assertEqual(lane["rss_kib"]["range"], [110, 130])

    def test_markdown_states_n_three_limit_and_no_causal_claim(self):
        for process in range(1, 4):
            self.write_receipt(process)
        with patch.object(uncertainty, "LANES", ("capture",)):
            markdown = uncertainty.to_markdown(uncertainty.summarize_results(self.results))
        self.assertIn("n=3", markdown)
        self.assertIn("not confidence intervals", markdown)
        self.assertIn("does not estimate a speedup", markdown)
        self.assertIn("causal effect", markdown)
        self.assertIn("within-process sample ranges", markdown)

    def test_missing_process_is_rejected(self):
        self.write_receipt(1)
        self.write_receipt(3)
        with patch.object(uncertainty, "LANES", ("capture",)):
            with self.assertRaisesRegex(ValueError, "exactly p1..p3"):
                uncertainty.summarize_results(self.results)

    def test_unknown_lane_receipts_are_rejected(self):
        for process in range(1, 4):
            self.write_receipt(process)
            self.write_receipt(process, lane="rogue")
        with patch.object(uncertainty, "LANES", ("capture",)):
            with self.assertRaisesRegex(ValueError, "unknown lane receipt"):
                uncertainty.summarize_results(self.results)

    def test_provenance_sidecars_are_not_process_receipts(self):
        for process in range(1, 4):
            self.write_receipt(process)
        (self.results / "source-provenance.json").write_text("{}")
        with patch.object(uncertainty, "LANES", ("capture",)):
            self.assertEqual(len(uncertainty.summarize_results(self.results)["lanes"]), 1)

    def test_mismatched_sample_counts_are_rejected(self):
        self.write_receipt(1)
        self.write_receipt(2, sample_count=21)
        self.write_receipt(3)
        with patch.object(uncertainty, "LANES", ("capture",)):
            with self.assertRaisesRegex(ValueError, "sample count changed"):
                uncertainty.summarize_results(self.results)

    def test_typed_refusal_gate_is_preserved(self):
        for process in range(1, 4):
            self.write_receipt(
                process,
                lane="limit_small",
                expected_success=False,
                error_class="caller_limit",
            )
        with patch.object(uncertainty, "LANES", ("limit_small",)):
            uncertainty.summarize_results(self.results)
            receipt = self.results / "limit_small-p2.json"
            payload = json.loads(receipt.read_text())
            payload["samples"][0]["error"]["class"] = "wrong_class"
            receipt.write_text(json.dumps(payload))
            with self.assertRaisesRegex(ValueError, "refusal class mismatch"):
                uncertainty.summarize_results(self.results)

    def test_refusal_lane_cannot_be_reported_as_success(self):
        for process in range(1, 4):
            self.write_receipt(
                process,
                lane="limit_small",
                expected_success=True,
            )
        with patch.object(uncertainty, "LANES", ("limit_small",)):
            with self.assertRaisesRegex(ValueError, "expected status does not match lane"):
                uncertainty.summarize_results(self.results)

    def test_success_lane_cannot_be_reported_as_refusal(self):
        for process in range(1, 4):
            self.write_receipt(process, expected_success=False)
        with patch.object(uncertainty, "LANES", ("capture",)):
            with self.assertRaisesRegex(ValueError, "expected status does not match lane"):
                uncertainty.summarize_results(self.results)

    def test_negative_measurement_input_and_rss_are_rejected(self):
        with patch.object(uncertainty, "LANES", ("capture",)):
            self.write_receipt(1)
            self.write_receipt(2)
            self.write_receipt(3)
            receipt = self.results / "capture-p1.json"
            payload = json.loads(receipt.read_text())
            payload["samples"][0]["elapsed_ns"] = -1
            receipt.write_text(json.dumps(payload))
            with self.assertRaisesRegex(ValueError, "elapsed_ns must be nonnegative"):
                uncertainty.summarize_results(self.results)

            for process in range(1, 4):
                self.write_receipt(process)
            receipt = self.results / "capture-p2.json"
            payload = json.loads(receipt.read_text())
            payload["input_bytes"] = -1
            receipt.write_text(json.dumps(payload))
            with self.assertRaisesRegex(ValueError, "input_bytes must be nonnegative"):
                uncertainty.summarize_results(self.results)

            for process in range(1, 4):
                self.write_receipt(process)
            timing = self.results / "capture-p3.time.txt"
            timing.write_text("Maximum resident set size (kbytes): -1\n")
            with self.assertRaisesRegex(ValueError, "rss must be nonnegative"):
                uncertainty.summarize_results(self.results)

    def test_malformed_input_digest_is_rejected(self):
        self.write_receipt(1)
        self.write_receipt(2)
        self.write_receipt(3)
        receipt = self.results / "capture-p3.json"
        payload = json.loads(receipt.read_text())
        payload["input_sha256"] = "not-a-digest"
        receipt.write_text(json.dumps(payload))
        with patch.object(uncertainty, "LANES", ("capture",)):
            with self.assertRaisesRegex(ValueError, "lowercase SHA-256"):
                uncertainty.summarize_results(self.results)


if __name__ == "__main__":
    unittest.main()
