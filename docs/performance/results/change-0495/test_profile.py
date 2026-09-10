#!/usr/bin/env python3
"""Focused parser and custody tests for the 0495 whole-child profiler.

These tests never build or launch a benchmark child.  Workload evidence is
created only by the coordinator after a retained after-normal build and a
fresh profile output directory have been authenticated.
"""

from __future__ import annotations

from pathlib import Path
import tempfile
import unittest
from unittest import mock
import sys

sys.path.insert(0, str(Path(__file__).resolve().parent))
import profile as profiler  # noqa: E402


class ProfileParserTests(unittest.TestCase):
    def test_perf_parser_separates_runtime_from_percentage(self) -> None:
        parsed = profiler.parse_perf_text(
            "100,,cycles,1120022288,81.00,,\n"
            "250,,instructions,1120022288,81.00,,\n"
            "10,,branches,1120022288,81.00,,\n"
            "2,,branch-misses,1120022288,81.00,,\n"
            "<not counted>,,LLC-loads,0,0.00,,\n"
            "# started on host\n"
        )
        cycles = parsed["events"]["cycles"]
        self.assertEqual(cycles["running_time_ns"], 1120022288)
        self.assertEqual(cycles["running_percent"], 81.0)
        self.assertNotEqual(cycles["running_percent"], 1120022288)
        self.assertEqual(parsed["events"]["LLC-loads"]["value"], None)
        self.assertEqual(parsed["derived"]["ipc"], 2.5)
        self.assertEqual(parsed["derived"]["branch_miss_rate"], 0.2)
        self.assertEqual(parsed["unparsed_lines"], [])
        self.assertEqual(len(parsed["ignored_comment_lines"]), 1)

    def test_strace_parser_accepts_blank_errors_column_and_checks_rows(self) -> None:
        parsed = profiler.parse_strace_text(
            "% time     seconds  usecs/call     calls    errors syscall\n"
            "------ ----------- ----------- --------- --------- ----------------\n"
            " 50.00    0.000010           5         2           0 read\n"
            " 25.00    0.000005           5         1 openat\n"
            " 25.00    0.000005           5         1           1 write\n"
            "------ ----------- ----------- --------- --------- ----------------\n"
            "100.00    0.000020           0         4           1 total\n"
        )
        self.assertEqual(parsed["total_calls"], 4)
        self.assertEqual(parsed["total_errors"], 1)
        self.assertIsNone(parsed["syscalls"]["openat"]["errors"])
        self.assertEqual(parsed["reported_total"]["calls"], 4)
        self.assertEqual(parsed["unparsed_lines"], [])
        self.assertIsNone(profiler._check_strace_parse(parsed))

    def test_strace_parser_rejects_unexplained_payload(self) -> None:
        parsed = profiler.parse_strace_text(
            "% time seconds usecs/call calls errors syscall\n"
            "not a syscall row\n"
        )
        self.assertEqual(len(parsed["unparsed_lines"]), 1)
        self.assertIn("unexplained", profiler._check_strace_parse(parsed) or "")

    def test_strace_total_accepts_blank_errors_column(self) -> None:
        parsed = profiler.parse_strace_text(
            "% time seconds usecs/call calls errors syscall\n"
            "100.00 0.000001 1 2 read\n"
            "100.00 0.000001 0 2 total\n"
        )
        self.assertEqual(parsed["reported_total"]["calls"], 2)
        self.assertIsNone(parsed["reported_total"]["errors"])
        self.assertIsNone(profiler._check_strace_parse(parsed))

    def test_provider_command_binds_api_and_canonical_arm_options(self) -> None:
        command = profiler._target_command(
            Path("/tmp/normal-provider"), "managed-api", "short", 3, 1,
            "a" * 40, Path("/tmp/report.json"), profiler.canonical_measure.CPU,
        )
        self.assertIn("--edit-api", command)
        self.assertIn("managed-api", command)
        self.assertIn("--provider", command)
        self.assertIn("short", command)
        self.assertIn("--max-range", command)
        self.assertIn("4096", command)
        self.assertIn("--trace-ranges", command)
        self.assertEqual(command[-1], "/tmp/report.json")

    def test_perf_record_command_omits_cpu_script_field(self) -> None:
        command = profiler._script_command(Path("/tmp/perf.data"), profiler.canonical_measure.CPU)
        self.assertIn("period", command[command.index("-F") + 1])
        self.assertNotIn(",cpu,", command[command.index("-F") + 1])
        record = profiler._record_command(
            Path("/tmp/normal-provider"), "unmanaged-api", 100, 3,
            "a" * 40, Path("/tmp/report.json"), profiler.canonical_measure.CPU,
            Path("/tmp/perf.data"), 199,
        )
        self.assertIn("--call-graph", record)
        self.assertIn("dwarf", record)
        self.assertIn("cycles:u", record)
        self.assertIn("100", record)
        self.assertIn("3", record)

    def test_perf_script_accepts_bare_symbols_and_ignores_headers(self) -> None:
        parsed = profiler.parse_perf_script_text(
            "# captured by perf\n"
            "litchi-perf-baseline 10 10 1.000: 7 cycles:u:\n"
            "      7f00 publish_docx_source_edit+0x1 (litchi)\n"
            "      7f00 run_sample+0x2 (litchi)\n"
            "      7f00 main+0x3 (litchi)\n"
            "\n"
            "litchi-perf-baseline 10 10 1.100: 1 cycles:u:\n"
            "      7f00 verify_docx_source_edit_output+0x1 (litchi)\n"
            "      7f00 run_sample+0x2 (litchi)\n"
        )
        self.assertEqual(parsed["sample_count"], 2)
        self.assertEqual(parsed["total_period"], 8)
        self.assertEqual(parsed["unparsed_lines"], [])
        summary = profiler.summarize_perf_stacks(parsed)
        classes = {row["class"]: row for row in summary["classes"]}
        self.assertEqual(classes["run_sample_publish_ancestor"]["period"], 7)
        self.assertEqual(classes["run_sample_output_oracle"]["period"], 1)
        self.assertEqual(summary["run_sample_period"], 8)
        self.assertEqual(summary["run_sample_publish_period"], 7)

    def test_perf_failure_does_not_hide_target_failure_as_unavailable(self) -> None:
        process = {"exit_code": 1, "timed_out": False}
        with tempfile.TemporaryDirectory() as directory:
            target_stderr = Path(directory) / "target.stderr"
            target_stderr.write_text("docx provider: permission denied\n", encoding="utf-8")
            status, reason = profiler._perf_failure_status(process, target_stderr)
            self.assertEqual(status, "failed")
            self.assertIn("nonzero", reason or "")
            pmu_stderr = Path(directory) / "pmu.stderr"
            pmu_stderr.write_text("No permission to enable cycles:u.\n", encoding="utf-8")
            status, _ = profiler._perf_failure_status(process, pmu_stderr)
            self.assertEqual(status, "unavailable")

    def test_strace_failure_distinguishes_ptrace_setup(self) -> None:
        process = {"exit_code": 1, "timed_out": False}
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "strace.stderr"
            path.write_text("strace: ptrace(PTRACE_TRACEME): Operation not permitted\n", encoding="utf-8")
            status, _ = profiler._strace_failure_status(process, path)
            self.assertEqual(status, "unavailable")
            path.write_text("docx provider failed: permission denied\n", encoding="utf-8")
            status, _ = profiler._strace_failure_status(process, path)
            self.assertEqual(status, "failed")

    def test_process_runner_refuses_existing_raw_paths(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            stdout, stderr = root / "stdout", root / "stderr"
            stdout.write_text("existing\n", encoding="utf-8")
            with self.assertRaises(FileExistsError):
                profiler._run_process(
                    ["/bin/true"], stdout=stdout, stderr=stderr,
                    env=dict(profiler.ENV), timeout=1,
                )

    def test_report_identity_passes_api_to_canonical_validator(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            report = Path(directory) / "report.json"
            report.write_text("{}\n", encoding="utf-8")
            expected = {
                "api": "managed-api",
                "provider": {"name": "owned", "kind": "owned"},
                "rows": [{"sample_index": 0}],
                "schema": profiler.canonical_measure.SCHEMA,
                "version": profiler.canonical_measure.VERSION,
                "source_archive_sha256": "a" * 64,
                "source_archive_bytes": 1,
                "source_revision": "b" * 40,
                "binary_sha256": "c" * 64,
                "binary_bytes": 123,
            }
            binary = Path(directory) / "provider"
            binary.write_bytes(b"provider")
            with mock.patch.object(profiler.canonical_measure, "validate_report", return_value=expected) as validator:
                result = profiler._provider_report_identity(
                    report, "owned", "managed-api", binary=binary,
                    revision="b" * 40, samples=1, warmup=1,
                )
            validator.assert_called_once_with(
                report, role="normal", api="managed-api", arm_name="owned",
                samples=1, warmups=1, source_revision="b" * 40,
                binary_sha256=profiler.sha(binary), binary_bytes=binary.stat().st_size,
            )
            self.assertEqual(result["api"], "managed-api")
            self.assertEqual(result["provider_reported"], "owned")


if __name__ == "__main__":
    raise SystemExit(unittest.main())
