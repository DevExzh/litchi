#!/usr/bin/env python3
"""Pure parser tests for the 0494 whole-child profiler."""

from __future__ import annotations

import json
from pathlib import Path
import sys
import tempfile
import unittest
from unittest import mock

sys.path.insert(0, str(Path(__file__).resolve().parent))
import profile as profiler  # noqa: E402


class ProfileParserTests(unittest.TestCase):
    def test_perf_parser_preserves_unavailable_and_derives_rates(self) -> None:
        parsed = profiler.parse_perf_text(
            "100,,cycles,100.00,,\n"
            "250,,instructions,100.00,,\n"
            "10,,branches,100.00,,\n"
            "2,,branch-misses,100.00,,\n"
            "<not counted>,,LLC-loads,0.00,,\n"
        )
        self.assertEqual(parsed["events"]["LLC-loads"]["value"], None)
        self.assertEqual(parsed["derived"]["ipc"], 2.5)
        self.assertEqual(parsed["derived"]["branch_miss_rate"], 0.2)

    def test_strace_parser_conserves_calls_and_errors(self) -> None:
        parsed = profiler.parse_strace_text(
            "% time     seconds  usecs/call     calls    errors syscall\n"
            "------ ----------- ----------- --------- --------- ----------------\n"
            " 50.00    0.000010           5         2           0 read\n"
            " 50.00    0.000010          10         1           1 openat\n"
            "------ ----------- ----------- --------- --------- ----------------\n"
            "100.00    0.000020           0         3           1 total\n"
        )
        self.assertEqual(parsed["total_calls"], 3)
        self.assertEqual(parsed["total_errors"], 1)
        self.assertEqual(parsed["syscalls"]["read"]["calls"], 2)
        self.assertEqual(parsed["reported_total"]["calls"], 3)

    def test_provider_arguments_are_explicit(self) -> None:
        self.assertEqual(profiler.provider_args("owned"), ["--provider", "owned"])
        self.assertEqual(profiler.provider_args("file-warm"), ["--provider", "file"])
        self.assertEqual(
            profiler.provider_args("short-read"),
            ["--provider", "short", "--short-range", "4096"],
        )
        with self.assertRaises(profiler.ProfileError):
            profiler.provider_args("remote")

    def test_perf_script_filters_publish_oracle_and_preflight(self) -> None:
        parsed = profiler.parse_perf_script_text(
            "# perf metadata\n"
            "docx-edit-provider 10 [002] 1.000: 7 cycles:u:\n"
            "      7f00 publish_docx_source_edit+0x1 (litchi)\n"
            "      7f00 docx_edit_provider::run_sample+0x2 (litchi)\n"
            "      7f00 main+0x3 (litchi)\n"
            "\n"
            "docx-edit-provider 10 [002] 1.100: cycles:u:\n"
            "      7f00 verify_docx_source_edit_output+0x1 (litchi)\n"
            "      7f00 docx_edit_provider::run_sample+0x2 (litchi)\n"
            "\n"
            "docx-edit-provider 10 [002] 1.200: 3 cycles:u:\n"
            "      7f00 docx_edit_provider::prepare+0x1 (litchi)\n"
        )
        self.assertEqual(parsed["sample_count"], 3)
        self.assertEqual(parsed["total_period"], 11)
        summary = profiler.summarize_perf_stacks(parsed)
        classes = {row["class"]: row for row in summary["classes"]}
        self.assertEqual(classes["run_sample_publish_ancestor"]["period"], 7)
        self.assertEqual(classes["run_sample_output_oracle"]["period"], 1)
        self.assertEqual(classes["preflight"]["period"], 3)
        self.assertEqual(summary["run_sample_period"], 8)
        self.assertEqual(summary["run_sample_publish_period"], 7)
        self.assertTrue(summary["folded_stacks"][0]["stack"].startswith("main;"))

    def test_perf_failure_is_not_unavailable_without_tool_diagnostic(self) -> None:
        process = {"exit_code": 1, "timed_out": False}
        with self.subTest("ordinary target failure"):
            status, reason = profiler._record_failure_status(process, Path("/definitely/missing"))
            self.assertEqual(status, "failed")
            self.assertIn("nonzero", reason or "")
        with self.subTest("PMU unavailable"):
            stderr = Path(self._testMethodName + ".stderr")
            try:
                stderr.write_text("No permission to enable cycles:u.\n", encoding="utf-8")
                status, _ = profiler._record_failure_status(process, stderr)
                self.assertEqual(status, "unavailable")
            finally:
                stderr.unlink(missing_ok=True)
        with self.subTest("target permission error remains failed"):
            stderr = Path(self._testMethodName + ".target.stderr")
            try:
                stderr.write_text("open input: permission denied\n", encoding="utf-8")
                status, _ = profiler._record_failure_status(process, stderr)
                self.assertEqual(status, "failed")
            finally:
                stderr.unlink(missing_ok=True)

    def test_perf_stat_retries_only_for_pmu_diagnostics(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            binary = root / "provider"
            binary.write_bytes(b"provider")
            binary.chmod(0o755)
            args = type("Args", (), {
                "samples": 1,
                "warmup": 0,
                "cpu": 2,
                "timeout": 10,
            })()

            def fake_process(argv, *, stdout, stderr, env, timeout):
                del argv, stdout, env, timeout
                stderr.write_text("No permission to enable cycles event.\n", encoding="utf-8")
                return {"exit_code": 1, "timed_out": False}

            with mock.patch.object(profiler, "_run_process", side_effect=fake_process):
                result = profiler._run_perf(binary, "owned", args, root, "a" * 40, {})
            self.assertEqual(result["status"], "unavailable")
            self.assertEqual(len(result["attempts"]), 2)

            def target_failure(argv, *, stdout, stderr, env, timeout):
                del argv, stdout, env, timeout
                stderr.write_text("target: permission denied\n", encoding="utf-8")
                return {"exit_code": 1, "timed_out": False}

            ordinary = root / "ordinary"
            ordinary.mkdir()
            with mock.patch.object(profiler, "_run_process", side_effect=target_failure):
                result = profiler._run_perf(binary, "owned", args, ordinary, "a" * 40, {})
            self.assertEqual(result["status"], "failed")
            self.assertEqual(len(result["attempts"]), 1)

    def test_owned_record_command_binds_requested_sampling(self) -> None:
        command = profiler._record_command(
            Path("/tmp/normal-provider"), 100, 3, "a" * 40,
            Path("/tmp/owned-report.json"), 2, Path("/tmp/perf.data"), 199,
        )
        self.assertIn("--call-graph", command)
        self.assertIn("dwarf", command)
        self.assertIn("-F", command)
        self.assertIn("199", command)
        self.assertIn("cycles:u", command)
        self.assertIn("100", command)
        self.assertIn("3", command)

    def test_report_identity_delegates_to_canonical_validator(self) -> None:
        revision = "b" * 40
        binary_sha256 = "c" * 64
        report = {
            "schema": profiler.REPORT_SCHEMA,
            "version": profiler.REPORT_VERSION,
            "case_name": profiler.REPORT_CASE_NAME,
            "provider": {"name": "owned", "kind": "owned"},
            "source_archive_sha256": profiler.EXPECTED_SOURCE_ARCHIVE_SHA256,
            "source_archive_bytes": profiler.EXPECTED_SOURCE_ARCHIVE_BYTES,
            "source_revision": revision,
            "binary_sha256": binary_sha256,
            "binary_bytes": 123,
            "rows": [{"sample_index": 0}, {"sample_index": 1}],
        }
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "report.json"
            path.write_text(json.dumps(report), encoding="utf-8")
            with mock.patch.object(
                profiler.canonical_measure,
                "validate_report",
                return_value=report,
            ) as validator:
                identity = profiler._report_identity(
                    path,
                    "owned",
                    expected_binary_sha256=binary_sha256,
                    expected_binary_bytes=123,
                    expected_source_revision=revision,
                    expected_samples=2,
                    expected_warmup=3,
                )
            validator.assert_called_once_with(
                path,
                role="normal",
                arm_name="owned",
                samples=2,
                warmups=3,
                source_revision=revision,
                binary_sha256=binary_sha256,
                binary_bytes=123,
            )
            self.assertTrue(identity["provider_match"])
            self.assertEqual(identity["rows"], 2)

    def test_report_identity_propagates_canonical_rejection(self) -> None:
        revision = "b" * 40
        binary_sha256 = "c" * 64
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "report.json"
            path.write_text("{}", encoding="utf-8")
            with mock.patch.object(
                profiler.canonical_measure,
                "validate_report",
                side_effect=profiler.canonical_measure.ProviderMatrixError("row invariant"),
            ):
                with self.assertRaises(profiler.ProfileError):
                    profiler._report_identity(
                        path,
                        "owned",
                        expected_binary_sha256=binary_sha256,
                        expected_binary_bytes=123,
                        expected_source_revision=revision,
                        expected_samples=2,
                        expected_warmup=3,
                    )


if __name__ == "__main__":
    raise SystemExit(unittest.main())
