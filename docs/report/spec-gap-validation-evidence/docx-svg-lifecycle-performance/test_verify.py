"""Tampered receipts must not qualify as performance evidence."""

import copy
import json
from pathlib import Path
import tempfile
import unittest

import verify


ROOT = Path(__file__).resolve().parent


def sample(lane):
    path = ROOT / "results" / f"smoke-{lane}-p1.json"
    return json.loads(path.read_text())["samples"][0]


class ReceiptVerificationTests(unittest.TestCase):
    def check(self, value, lane="single_attach_1"):
        verify.verify_sample(value, lane, lane not in verify.EXPECTED_REFUSALS, Path("test.json"))

    def test_retained_success_and_refusal_are_valid(self):
        for lane in ("single_attach_1", "batch_attach_64"):
            self.check(sample(lane), lane)

    def test_refusal_requires_typed_match_and_complete_source_readback(self):
        baseline = sample("batch_attach_64")
        for key in ("source_readback_physical_ok", "source_readback_metadata_ok"):
            with self.subTest(key=key):
                altered = copy.deepcopy(baseline)
                altered[key] = False
                with self.assertRaises(AssertionError):
                    self.check(altered, "batch_attach_64")
        altered = copy.deepcopy(baseline)
        altered["error"]["typed_match"] = False
        with self.assertRaises(AssertionError):
            self.check(altered, "batch_attach_64")

    def test_numeric_metrics_are_unsigned_integers(self):
        baseline = sample("single_attach_1")
        for key in ("allocation_calls", "deallocation_calls", "elapsed_ns", "peak_live_delta"):
            for invalid in (-1, True, "1", 1.5):
                with self.subTest(key=key, invalid=invalid):
                    altered = copy.deepcopy(baseline)
                    altered[key] = invalid
                    with self.assertRaises(AssertionError):
                        self.check(altered)

    def test_named_phases_fit_the_measured_interval(self):
        baseline = sample("single_attach_1")
        for invalid in (-1, True, baseline["elapsed_ns"] + 1):
            altered = copy.deepcopy(baseline)
            altered["phases"]["capture_ns"] = invalid
            with self.subTest(invalid=invalid), self.assertRaises(AssertionError):
                self.check(altered)

    def test_rss_receipt_requires_one_positive_integer(self):
        with tempfile.TemporaryDirectory(prefix="litchi-docx-rss-test-") as directory:
            timing = Path(directory) / "process.time.txt"
            stderr = Path(directory) / "process.stderr.log"
            stderr.write_bytes(b"")
            marker = "Maximum resident set size (kbytes): "
            timing.write_text(marker + "1024\nExit status: 0\n")
            verify.verify_timing(timing, stderr)
            for value in ("-1", "0", "1.5", "unknown", "1\n" + marker + "2"):
                timing.write_text(marker + value + "\nExit status: 0\n")
                with self.subTest(value=value), self.assertRaises(AssertionError):
                    verify.verify_timing(timing, stderr)


if __name__ == "__main__":
    unittest.main()
