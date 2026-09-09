"""Focused pure-Python checks for the input metadata profile driver."""

import gzip
import json
from pathlib import Path
import sys
import tempfile
import unittest
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parent))
import measure_routes as routes
import profile_input_metadata as input_metadata
import profile_routes


class InputMetadataProfileTests(unittest.TestCase):
    def test_summary_command_contains_metadata_filter_and_exact_route_argv(self):
        route_argv = [
            "/usr/bin/time",
            "-v",
            "-o",
            "/tmp/resource.txt",
            "/usr/bin/taskset",
            "-c",
            "2",
            "/tmp/docx",
            "--json",
            "/tmp/report.json",
        ]
        command = input_metadata._strace_command(
            {"path": "/usr/bin/strace", "status": "available"},
            route_argv,
            Path("/tmp/summary.txt"),
            summary=True,
        )
        self.assertIn("-f", command)
        self.assertIn("-c", command)
        self.assertIn(input_metadata.TRACE_FILTER, command)
        self.assertEqual(command[command.index("--") + 1:], route_argv)
        for syscall in ("fstat", "newfstatat", "statx", "pread64", "read", "write", "open", "close", "lseek", "sync"):
            self.assertIn(syscall, input_metadata.TRACE_FILTER)

    def test_owned_and_file_arms_bind_deterministic_inputs(self):
        workload = "s64-a16384-short-c64"
        owned = input_metadata._arm_for(workload, "owned")
        file_arm = input_metadata._arm_for(workload, "file")
        self.assertEqual(owned["provider"], "deterministic")
        self.assertEqual(owned["input_mode"], "owned")
        self.assertIsNone(owned["input_file"])
        self.assertEqual(file_arm["provider"], "deterministic")
        self.assertEqual(file_arm["input_mode"], "file")
        self.assertTrue(file_arm["input_file"].endswith("-source.docx"))
        file_binding = input_metadata._input_metadata(file_arm)
        self.assertEqual(file_binding["mode"], "file")
        self.assertGreater(file_binding["bytes"], 0)
        self.assertEqual(
            file_binding["sha256"],
            routes.meta(Path(file_binding["absolute_path"]))["sha256"],
        )

    def test_raw_trace_hash_is_recorded_before_gzip(self):
        with tempfile.TemporaryDirectory() as temporary:
            raw = Path(temporary) / "strace-raw.log"
            raw.write_bytes(b"openat(3, \"source.docx\", O_RDONLY) = 4\\n")
            binding = input_metadata._gzip_after_hash(raw)
            self.assertEqual(binding["raw_before_compression"]["sha256"], routes.meta(raw)["sha256"])
            compressed = raw.with_name(raw.name + ".gz")
            with gzip.open(compressed, "rb") as stream:
                self.assertEqual(stream.read(), raw.read_bytes())
            self.assertTrue(binding["compressed"]["present"])

    def test_unlaunchable_strace_is_unavailable(self):
        status, failure, error = input_metadata._classify_process(
            {
                "returncode": None,
                "timed_out": False,
                "route_started": False,
                "launch_error_type": "FileNotFoundError",
                "launch_error": "FileNotFoundError: strace",
            },
            "",
        )
        self.assertEqual((status, failure), ("unavailable", "profiler_unavailable"))
        self.assertIn("strace", error or "")

    def test_success_requires_core_report_resource_and_summary(self):
        with tempfile.TemporaryDirectory() as temporary:
            destination = Path(temporary) / "output"
            destination.mkdir()
            arm = input_metadata._arm_for("s64-a64-short-c64", "owned")
            binary = {
                "path": "/tmp/docx_replayable_tail_append",
                "bytes": 1,
                "sha256": "a" * 64,
                "executable": True,
            }
            build = {"path": "route-attempts/formal1/build-normal.json", "sha256": "b" * 64}

            def successful_run(command, *, stdout_path, stderr_path, timeout_seconds):
                stdout_path.write_bytes(b"stdout")
                stderr_path.write_bytes(b"")
                directory = stdout_path.parent
                (directory / "report.json").write_text("{}", encoding="utf-8")
                (directory / "resource.txt").write_text("Maximum resident set size (kbytes): 1\\n", encoding="utf-8")
                (directory / "strace-summary.txt").write_text("% time seconds calls syscall\\n", encoding="utf-8")
                return {"returncode": 0, "timed_out": False, "route_started": True}

            with (
                patch.object(input_metadata.profile_driver, "_run_command", side_effect=successful_run),
                patch.object(input_metadata.routes, "_check_axis_report", return_value={}),
            ):
                result = input_metadata._run_trace(
                    destination=destination,
                    binary=binary,
                    build=build,
                    protocol_sha256="c" * 64,
                    tool_info={"path": "/usr/bin/strace", "status": "available"},
                    attempt="metadata1",
                    build_attempt="formal1",
                    case_label="s64-a64-short-c64",
                    arm=arm,
                    input_metadata=input_metadata._input_metadata(arm),
                    timeout_seconds=1,
                    raw=False,
                )
            self.assertEqual(result["status"], "ok")
            receipt_path = destination / result["label"] / "receipt.json"
            receipt = json.loads(receipt_path.read_text(encoding="utf-8"))
            self.assertEqual(receipt["validation_status"], "ok")
            self.assertTrue(receipt["profile_artifact"]["present"])


if __name__ == "__main__":
    unittest.main()
