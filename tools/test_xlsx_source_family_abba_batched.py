#!/usr/bin/env python3
"""Focused contract tests for the XLSX source-family evidence adapter."""

from __future__ import annotations

import importlib.util
from pathlib import Path
import sys
import tempfile
import unittest
from unittest import mock
from contextlib import redirect_stderr
import io


MODULE_PATH = Path(__file__).with_name("xlsx_source_family_abba_batched.py")
SPEC = importlib.util.spec_from_file_location("xlsx_source_family_abba_batched", MODULE_PATH)
assert SPEC is not None and SPEC.loader is not None
ADAPTER = importlib.util.module_from_spec(SPEC)
sys.modules[SPEC.name] = ADAPTER
SPEC.loader.exec_module(ADAPTER)


def child_report(case: str, elapsed: list[int], read_calls: list[int], read_bytes: list[int]) -> dict:
    samples = len(elapsed)
    return {
        "results": [{
            "case": case,
            "elapsed_ns": {"samples": elapsed},
            "output_sha256": "same-output",
            "sink": {
                "accepted_bytes": [1000] * samples,
                "largest_write": [4096] * samples,
            },
            "source": {
                "read_calls": read_calls,
                "read_bytes": read_bytes,
            },
        }],
    }


def event(revision: str, elapsed: list[int], read_calls: list[int], read_bytes: list[int]) -> dict:
    result = ADAPTER.extract_result(child_report("case", elapsed, read_calls, read_bytes), "case", len(elapsed))
    source = ADAPTER.source_signature(result, len(elapsed))
    sink = ADAPTER.sink_signature(result, len(elapsed))
    summary = ADAPTER.elapsed_summary(elapsed, len(elapsed))
    return {
        "revision": revision,
        "output_sha256": "same-output",
        "source": source,
        "sink": sink,
        **summary,
        "per_op_elapsed_ns": list(elapsed),
    }


class XlsxSourceFamilyAdapterTests(unittest.TestCase):
    def test_elapsed_samples_are_preserved_and_mean_is_normalized_once(self) -> None:
        summary = ADAPTER.elapsed_summary([100, 120, 140], 3)
        self.assertEqual(summary["sum_elapsed_ns"], 360)
        self.assertEqual(summary["mean_elapsed_ns"], 120)
        self.assertEqual(summary["per_op_elapsed_ns"], [100, 120, 140])

    def test_real_shaped_vectors_allow_different_endpoint_n(self) -> None:
        before = event("before", [100, 120], [7, 7], [70, 70])
        after_one = event("after", [220, 240, 260], [9, 9, 9], [90, 90, 90])
        after_two = event("after", [200, 220, 240], [9, 9, 9], [90, 90, 90])
        before_two = event("before", [110, 130], [7, 7], [70, 70])
        ADAPTER.validate_family_observations(
            "page_margin", "source", [before, after_one, after_two, before_two]
        )
        legs = [
            {"revision": "before", "mean_elapsed_ns": before["mean_elapsed_ns"]},
            {"revision": "after", "mean_elapsed_ns": after_one["mean_elapsed_ns"]},
            {"revision": "after", "mean_elapsed_ns": after_two["mean_elapsed_ns"]},
            {"revision": "before", "mean_elapsed_ns": before_two["mean_elapsed_ns"]},
        ]
        # Before mean=(110+120)/2, after mean=(240+220)/2.
        self.assertAlmostEqual(ADAPTER.paired_block_mean_ratio(legs), 230 / 115)

    def test_source_and_sink_change_between_revisions_but_must_stay_stable_within_one(self) -> None:
        before = event("before", [100, 110], [7, 7], [70, 70])
        changed_before = event("before", [105, 115], [8, 8], [70, 70])
        with self.assertRaises(ADAPTER.RunFailure):
            ADAPTER.validate_family_observations("page_margin", "source", [before, changed_before])
        missing_sink = dict(before)
        missing_sink["sink"] = {"accepted_bytes": [1000, 1000]}
        with self.assertRaises(ADAPTER.RunFailure):
            ADAPTER.validate_family_observations("page_margin", "source", [missing_sink])

    def test_elapsed_and_source_vector_lengths_are_checked(self) -> None:
        with self.assertRaises(ADAPTER.RunFailure):
            ADAPTER.extract_result(
                child_report("case", [100], [7], [70]), "case", 2
            )
        with self.assertRaises(ADAPTER.RunFailure):
            result = ADAPTER.extract_result(
                child_report("case", [100, 120], [7, 7], [70, 70]), "case", 2
            )
            ADAPTER.source_signature({**result, "source": {"read_calls": [7], "read_bytes": [70, 70]}}, 2)

    def test_result_directory_must_be_empty(self) -> None:
        with tempfile.TemporaryDirectory() as root:
            out = Path(root) / "evidence"
            out.mkdir()
            (out / "old.json").write_text("old\n", encoding="utf-8")
            with self.assertRaises(ADAPTER.RunFailure):
                ADAPTER.Recorder(out)

    def test_preflight_only_is_not_a_second_name_for_plan(self) -> None:
        argv = [
            str(MODULE_PATH), "--before-bin", "before", "--after-bin", "after",
            "--out-dir", "out", "--preflight-only",
        ]
        with mock.patch.object(sys, "argv", argv):
            with redirect_stderr(io.StringIO()), self.assertRaises(SystemExit):
                ADAPTER.parse_args()

    def test_quiet_scan_recognizes_benchmark_children(self) -> None:
        for executable in (
            "litchi-perf-baseline", "litchi-perf-baseline-v2", "xlsx_probe",
            "scoped_patch_probe", "snappy-probe-v1", "ole_w5_probe",
        ):
            self.assertTrue(ADAPTER.is_heavy_process(executable, executable))
        self.assertTrue(ADAPTER.is_heavy_process("python3", "python3 /tmp/snappy-probe.py"))
        self.assertFalse(ADAPTER.is_heavy_process("python3", "python3 adapter.py"))


if __name__ == "__main__":
    unittest.main()
