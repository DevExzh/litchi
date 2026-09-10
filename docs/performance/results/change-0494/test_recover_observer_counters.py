#!/usr/bin/env python3
"""Focused tests for retained observer-counter and stack corrections."""

from __future__ import annotations

from pathlib import Path
import sys
import unittest

sys.path.insert(0, str(Path(__file__).resolve().parent))
import recover_observer_counters as correction  # noqa: E402


class ObserverCorrectionTests(unittest.TestCase):
    def test_perf_csv_uses_field_three_for_time_and_field_four_for_percent(self) -> None:
        parsed = correction.parse_perf_stat_corrected(
            "# perf stat header\n"
            "5711648969,,cycles,1158723044,83.00,,\n"
        )
        cycles = parsed["events"]["cycles"]
        self.assertEqual(cycles["running_time_ns"], 1158723044)
        self.assertEqual(cycles["running_percent"], 83.0)
        self.assertNotEqual(cycles["running_percent"], 1158723044)
        self.assertEqual(parsed["ignored_comment_lines"], ["# perf stat header"])

    def test_perf_csv_rejects_unexplained_payload(self) -> None:
        with self.assertRaises(correction.CorrectionError):
            correction.parse_perf_stat_corrected("bad payload\n")

    def test_strace_accepts_implicit_zero_errors_and_conserves_totals(self) -> None:
        parsed = correction.parse_strace_corrected(
            "% time seconds usecs/call calls errors syscall\n"
            " 50.00 0.001 1 2 read\n"
            " 25.00 0.000 2 1 1 openat\n"
            "------ ----------- --------- --------- --------- ------------------\n"
            "100.00 0.001 1 3 1 total\n"
        )
        self.assertEqual(parsed["total_calls"], 3)
        self.assertEqual(parsed["total_errors"], 1)
        self.assertEqual(parsed["syscalls"]["read"]["errors"], 0)
        self.assertEqual(parsed["syscalls"]["read"]["error_source"], "implicit-zero")

    def test_strace_rejects_unexplained_rows_or_total_mismatch(self) -> None:
        with self.assertRaises(correction.CorrectionError):
            correction.parse_strace_corrected(
                " 50.00 0.001 1 2 read\n"
                "100.00 0.001 1 9 0 total\n"
            )

    def test_bare_run_sample_is_classified_but_marked_ambiguous(self) -> None:
        classification, qualification = correction.classify_stack_corrected(
            ["publish_docx_source_edit<&mut", "run_sample"]
        )
        self.assertEqual(classification, "run_sample_publish_ancestor")
        self.assertEqual(qualification, "bare_run_sample_ambiguous")

    def test_stack_summary_counts_bare_symbols(self) -> None:
        summary = correction.summarize_stacks_corrected({
            "samples": [
                {"period": 7, "frames": ["publish_docx_source_edit<&mut", "run_sample"]},
                {"period": 5, "frames": ["verify_docx_source_edit_output", "run_sample"]},
            ],
            "unparsed_lines": [],
        })
        self.assertEqual(summary["sample_count"], 2)
        self.assertEqual(summary["run_sample_period"], 12)
        self.assertEqual(summary["run_sample_publish_period"], 7)
        self.assertTrue(summary["ambiguous_bare_run_sample"])

    def test_stack_comments_are_explicit_and_payload_is_rejected(self) -> None:
        summary = correction.summarize_stacks_corrected({
            "samples": [{"period": 1, "frames": ["run_sample"]}],
            "unparsed_lines": ["# perf header"],
        })
        self.assertEqual(summary["ignored_comment_count"], 1)
        with self.assertRaises(correction.CorrectionError):
            correction.summarize_stacks_corrected({
                "samples": [{"period": 1, "frames": ["run_sample"]}],
                "unparsed_lines": ["malformed payload"],
            })

    def test_recovery_r3_binds_archived_recovery_helper(self) -> None:
        archive = correction._recovery_helper_archive(
            correction.ROOT / "profiling-recovery-r3-helper-sources"
        )
        self.assertEqual(
            archive["archived"]["recover_profile.py"]["sha256"],
            "bea4eac6370dfe1a704321a9f829283182735e9646cb34b5981149006c9c5598",
        )
        self.assertFalse(archive["current_matches_archived"]["recover_profile.py"])


if __name__ == "__main__":
    raise SystemExit(unittest.main())
