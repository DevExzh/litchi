"""Unit tests for the standard-library resource-profile parser."""

from __future__ import annotations

import json
import tempfile
import unittest
from pathlib import Path

from tools import perf_resource_profile


class ResourceProfileParserTests(unittest.TestCase):
    def test_time_parser_extracts_rss_and_elapsed(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "time.txt"
            path.write_text(
                """\
Maximum resident set size (kbytes): 12345
User time (seconds): 1.25
System time (seconds): 0.50
\tElapsed (wall clock) time (h:mm:ss or m:ss): 0:02.75
Voluntary context switches: 12
Minor (reclaiming a frame) page faults: 34
""",
                encoding="utf-8",
            )
            parsed = perf_resource_profile.parse_time_report(path)
        self.assertEqual(parsed["status"], "ok")
        self.assertEqual(parsed["max_rss_kib"], 12345)
        self.assertEqual(parsed["user_seconds"], 1.25)
        self.assertEqual(parsed["elapsed_wall_seconds"], 2.75)
        self.assertEqual(parsed["voluntary_context_switches"], 12)

    def test_strace_parser_buckets_successes_and_failures(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "strace.log"
            path.write_text(
                """\
1 read(3, \"x\", 10) = 10
1 read(3, \"\", 0) = 0
1 read(3, 0x0, 4) = -1 EIO (Input/output error)
1 write(1, \"x\", 4096) = 4096
1 write(1, \"x\", 70000) = 70000
""",
                encoding="utf-8",
            )
            parsed = perf_resource_profile.parse_strace(path, 0)
        self.assertEqual(parsed["status"], "ok")
        self.assertEqual(parsed["syscalls"]["read"]["calls"], 2)
        self.assertEqual(parsed["syscalls"]["read"]["failed_calls"], 1)
        self.assertEqual(parsed["syscalls"]["read"]["returned_bytes"], 10)
        self.assertEqual(parsed["syscalls"]["write"]["calls"], 2)
        self.assertEqual(parsed["syscalls"]["write"]["returned_bytes"], 74096)
        self.assertEqual(parsed["syscalls"]["write"]["size_buckets"]["4096-16383"], 1)
        self.assertEqual(parsed["syscalls"]["write"]["size_buckets"]["65536-262143"], 1)

    def test_perf_parser_fail_closed_with_null_counters(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "perf.txt"
            stderr = Path(directory) / "stderr.txt"
            path.write_text("<not counted>,,cycles\n", encoding="utf-8")
            stderr.write_text("No permission to enable cycles event.\n", encoding="utf-8")
            parsed = perf_resource_profile.parse_perf_stat(path, 255, stderr)
        self.assertEqual(parsed["status"], "unsupported")
        self.assertIsNone(parsed["counters"]["cycles"]["value"])
        self.assertFalse(parsed["counters"]["cycles"]["available"])

    def test_scaling_analysis_is_explicit_and_finite(self):
        measurements = [
            {
                "elapsed_ns": {"p50": 1000},
                "execution": {"worker_count": 1, "logical_tasks": 16, "logical_bytes": 100},
            },
            {
                "elapsed_ns": {"p50": 600},
                "execution": {"worker_count": 2, "logical_tasks": 16, "logical_bytes": 100},
            },
            {
                "elapsed_ns": {"p50": 400},
                "execution": {"worker_count": 4, "logical_tasks": 16, "logical_bytes": 100},
            },
        ]
        summary = perf_resource_profile.scaling_analysis(measurements)
        self.assertEqual(summary["status"], "observed")
        self.assertEqual(summary["baseline_workers"], 1)
        self.assertEqual(summary["rows"][1]["speedup_vs_baseline"], 1000 / 600)
        self.assertTrue(summary["rows"][1]["amdahl_serial_fraction_valid"])
        self.assertIsInstance(summary["rows"][1]["amdahl_serial_fraction"], float)
        self.assertTrue(summary["classification"])
        json.dumps(summary, allow_nan=False)

    def test_scaling_analysis_nulls_invalid_amdahl_fraction_before_classifying(self):
        measurements = [
            {"elapsed_ns": {"p50": 1000}, "execution": {"worker_count": 1}},
            {"elapsed_ns": {"p50": 2000}, "execution": {"worker_count": 2}},
        ]
        summary = perf_resource_profile.scaling_analysis(measurements)
        row = summary["rows"][1]
        self.assertIsNone(row["amdahl_serial_fraction"])
        self.assertGreater(row["amdahl_serial_fraction_raw"], 1.0)
        self.assertFalse(row["amdahl_serial_fraction_valid"])
        self.assertEqual(row["amdahl_serial_fraction_reason"], "outside_0_1")
        self.assertEqual(summary["classification"], "nonideal_or_measurement_noise")
        json.dumps(summary, allow_nan=False)

    def test_prebuilt_identity_is_hash_only_and_requires_clean_build_rerun(self):
        with tempfile.TemporaryDirectory() as directory:
            binary = Path(directory) / "litchi-perf-baseline"
            binary.write_bytes(b"prebuilt")
            identity = perf_resource_profile.build_identity(
                binary,
                build_command=["cargo", "build", "--locked"],
                build_result={"executed": False},
                source_identity={"revision": "head", "head_tree": "tree"},
                rerun_command=["python3", "tools/perf_resource_profile.py", "run", "--build"],
            )
        self.assertEqual(identity["provenance"]["status"], "prebuilt_binary_hash_only")
        self.assertTrue(identity["provenance"]["rerun_required"])
        self.assertEqual(identity["source_content_identity"]["revision"], "head")
        self.assertEqual(identity["build_result"]["executed"], False)
        json.dumps(identity, allow_nan=False)

    def test_logical_measurements_retain_identity_and_counters(self):
        report = {
            "results": [
                {
                    "case": "rtf_streaming_create",
                    "corpus": {
                        "name": "rtf-medium",
                        "generator": "rtf-v1",
                        "archive_sha256": "a" * 64,
                        "archive_bytes": 123,
                    },
                    "elapsed_ns": {"unit": "ns", "p50": 10, "p95": 11, "p99": 12},
                    "sink": {"accepted_bytes": 123, "write_calls": 4},
                }
            ]
        }
        rows = perf_resource_profile.logical_measurements(report)
        self.assertEqual(rows[0]["case"], "rtf_streaming_create")
        self.assertEqual(rows[0]["corpus"]["archive_sha256"], "a" * 64)
        self.assertEqual(rows[0]["sink"]["write_calls"], 4)
        self.assertNotIn("target_payload_sha256", rows[0]["corpus"])

    def test_missing_heaptrack_is_not_zero(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "missing.txt"
            parsed = perf_resource_profile.parse_heaptrack_print(path)
        self.assertEqual(parsed["status"], "missing")
        self.assertNotIn("peak_heap_bytes", parsed)

    def test_heaptrack_summary_and_histogram_are_parsed(self):
        with tempfile.TemporaryDirectory() as directory:
            summary = Path(directory) / "print.txt"
            histogram = Path(directory) / "hist.tsv"
            summary.write_text(
                """\
calls to allocation functions: 12 (3/s)
34 temporary allocations of 56 allocations in total (60.7%)
peak heap memory consumption: 2.50M
peak RSS (including heaptrack overhead): 4.00M
""",
                encoding="utf-8",
            )
            histogram.write_text("4\t3\n8\t2\n", encoding="utf-8")
            parsed = perf_resource_profile.parse_heaptrack_print(summary)
            allocated = perf_resource_profile.parse_heaptrack_histogram(histogram)
        self.assertEqual(parsed["allocation_calls"], 12)
        self.assertEqual(parsed["temporary_allocations"], 34)
        self.assertEqual(parsed["peak_heap_bytes"], int(2.5 * 1024 * 1024))
        self.assertEqual(parsed["peak_rss_bytes"], 4 * 1024 * 1024)
        self.assertEqual(allocated, 28)

    def test_filesystem_evidence_omits_raw_request_arrays(self):
        compact = perf_resource_profile.compact_filesystem_evidence(
            [
                {
                    "case": "cfb-save",
                    "samples": [
                        {
                            "logical_read_calls": 4,
                            "logical_read_request_sizes": [1, 2, 3, 4],
                            "logical_read_request_size_buckets": {"1-511": 4},
                        }
                    ],
                }
            ]
        )
        sample = compact[0]["samples"][0]
        self.assertNotIn("logical_read_request_sizes", sample)
        self.assertEqual(sample["logical_read_calls"], 4)
        self.assertEqual(sample["logical_read_request_size_buckets"]["1-511"], 4)


if __name__ == "__main__":
    unittest.main()
