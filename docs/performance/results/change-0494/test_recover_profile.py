#!/usr/bin/env python3
"""Focused tests for retained perf-script recovery."""

from __future__ import annotations

from pathlib import Path
import sys
import unittest

sys.path.insert(0, str(Path(__file__).resolve().parent))
import profile as profiler  # noqa: E402
import recover_profile as recovery  # noqa: E402


class RecoveryTests(unittest.TestCase):
    def test_script_command_omits_cpu_field(self) -> None:
        command = recovery.script_command(Path("/tmp/perf.data"), 2)
        self.assertEqual(command[command.index("-F") + 1], recovery.SCRIPT_FIELDS)
        self.assertNotIn("cpu", recovery.SCRIPT_FIELDS.split(","))
        self.assertEqual(command[-2:], ["-i", "/tmp/perf.data"])

    def test_canonical_parser_accepts_header_without_cpu(self) -> None:
        parsed = profiler.parse_perf_script_text(
            "litchi-perf-baseline 10 10 1.000: 7 cycles:u:\n"
            "      7f00 publish_docx_source_edit+0x1 (litchi)\n"
            "      7f00 docx_edit_provider::run_sample+0x2 (litchi)\n"
            "\n"
        )
        self.assertEqual(parsed["sample_count"], 1)
        self.assertEqual(parsed["total_period"], 7)
        summary = profiler.summarize_perf_stacks(parsed)
        self.assertEqual(summary["run_sample_publish_period"], 7)

    def test_recovery_label_is_separate_from_original_attempt(self) -> None:
        self.assertEqual(recovery.RECOVERY_LABEL, "owned-perf-script-no-cpu")
        self.assertNotEqual(recovery.RECOVERY_LABEL, recovery.PROFILE_ATTEMPT)

    def test_capture_archive_binds_historical_helper_and_records_future_drift(self) -> None:
        archive = recovery._helper_archive(
            recovery.ROOT / "profiling-r1-helper-sources"
        )
        self.assertEqual(
            archive["profile"]["sha256"],
            "3b7e6a067ad3ca3020218fc3fc8dd2e146ffc34518ee82bd71fcfded9ecf007b",
        )
        self.assertFalse(archive["current_matches_archived"]["profile.py"])
        self.assertTrue(archive["current_matches_archived"]["test_profile.py"])


if __name__ == "__main__":
    raise SystemExit(unittest.main())
