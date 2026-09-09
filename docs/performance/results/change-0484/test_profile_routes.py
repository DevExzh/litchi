"""Focused pure-Python checks for the external route profiler driver."""

from pathlib import Path
import sys
import tempfile
import unittest
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parent))
import measure_routes as routes
import profile_routes


class ProfileRouteHelperTests(unittest.TestCase):
    def test_selection_normalizes_aliases_and_deduplicates(self):
        self.assertEqual(
            profile_routes._normalize_selection(
                ["perf_stat", "perf-stat", "perf_record"],
                profile_routes.PROFILE_TOOLS,
                profile_routes.PROFILE_TOOLS,
                label="profiler",
            ),
            ("perf-stat", "perf-record"),
        )

    def test_command_wrappers_keep_route_argv_and_scope(self):
        with tempfile.TemporaryDirectory() as temporary:
            directory = Path(temporary)
            route_argv = [
                "/usr/bin/time", "-v", "-o", "/tmp/resource.txt",
                "/usr/bin/taskset", "-c", "2", "/tmp/binary", "--json", "/tmp/report.json",
            ]
            fake = {"path": "/usr/bin/perf", "status": "available"}
            stat = profile_routes._command_for("perf-stat", fake, route_argv, directory)
            record = profile_routes._command_for("perf-record", fake, route_argv, directory)
            self.assertEqual(stat[0:2], ["/usr/bin/perf", "stat"])
            self.assertIn("--", stat)
            self.assertEqual(stat[stat.index("--") + 1:], route_argv)
            self.assertEqual(record[0:2], ["/usr/bin/perf", "record"])
            self.assertIn(str(directory / "perf.data"), record)

            trace = profile_routes._command_for(
                "strace", {"path": "/usr/bin/strace", "status": "available"}, route_argv, directory
            )
            self.assertIn(
                "trace=read,write,pread64,pwrite64,readv,writev,lseek,openat,close,fsync,fdatasync,unlink,unlinkat",
                trace,
            )
            self.assertEqual(trace[trace.index("--") + 1:], route_argv)

            heaptrack = profile_routes._command_for(
                "heaptrack", {"path": "/usr/bin/heaptrack", "status": "available"}, route_argv, directory
            )
            self.assertEqual(heaptrack[:7], route_argv[:7])
            self.assertEqual(heaptrack[7], "/usr/bin/heaptrack")
            self.assertEqual(heaptrack[-2:], route_argv[-2:])

    def test_perf_missing_counter_is_null_and_never_zero(self):
        with tempfile.TemporaryDirectory() as temporary:
            path = Path(temporary) / "perf-stat.txt"
            path.write_text(
                "123,cycles,cycles,100.00\n"
                "<not supported>,,instructions,\n"
                "0,counts,page-faults,100.00\n",
                encoding="utf-8",
            )
            counters = profile_routes._parse_perf_stat(path)
            self.assertEqual(counters["cycles"]["value"], 123)
            self.assertIsNone(counters["instructions"]["value"])
            self.assertFalse(counters["instructions"]["available"])
            self.assertEqual(counters["page-faults"]["value"], 0)
            self.assertTrue(counters["page-faults"]["available"])

    def test_required_profile_artifacts_include_nonempty_core_outputs(self):
        with tempfile.TemporaryDirectory() as temporary:
            directory = Path(temporary)
            report = directory / "report.json"
            resource = directory / "resource.txt"
            raw = directory / "perf-stat.txt"
            report.write_text("{}", encoding="utf-8")
            resource.write_bytes(b"")
            raw.write_text("0,0,cycles,100.00\n", encoding="utf-8")
            missing = profile_routes._missing_profile_artifacts(
                "perf-stat", directory, report, resource
            )
            self.assertTrue(any(item.endswith("resource.txt") for item in missing))
            self.assertFalse(any(item.endswith("report.json") for item in missing))
            self.assertFalse(any(item.endswith("perf-stat.txt") for item in missing))

    def test_failure_classification_distinguishes_permission_and_unsupported(self):
        self.assertEqual(
            profile_routes._classify_failure(
                "perf-stat", 255, "No permission to enable cycles event."
            ),
            "permission_denied",
        )
        self.assertEqual(
            profile_routes._classify_failure("perf-record", 1, "event syntax error"),
            "unsupported",
        )
        self.assertEqual(profile_routes._classify_failure("strace", 1, "bad trace"), "failed")
        self.assertEqual(profile_routes._classify_failure("strace", None, "", timed_out=True), "timeout")

    def test_timeout_terminates_the_owned_process_group(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            result = profile_routes._run_command(
                [sys.executable, "-c", "import time; time.sleep(2)"],
                stdout_path=root / "stdout.txt",
                stderr_path=root / "stderr.txt",
                timeout_seconds=0.02,
            )
            self.assertTrue(result["timed_out"])
            self.assertTrue(result["process_group_termination"]["start_new_session"])
            self.assertTrue((root / "stdout.txt").is_file())
            self.assertTrue((root / "stderr.txt").is_file())

    def test_timeout_kills_descendant_that_exits_parent_on_sigterm(self):
        tree = (
            "import os, signal, time\n"
            "child = os.fork()\n"
            "if child == 0:\n"
            "    signal.signal(signal.SIGTERM, signal.SIG_IGN)\n"
            "    time.sleep(5)\n"
            "else:\n"
            "    print(child, flush=True)\n"
            "    time.sleep(5)\n"
        )
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            result = profile_routes._run_command(
                [sys.executable, "-c", tree],
                stdout_path=root / "stdout.txt",
                stderr_path=root / "stderr.txt",
                timeout_seconds=0.02,
            )
            self.assertTrue(result["timed_out"])
            termination = result["process_group_termination"]
            self.assertIn("SIGTERM", termination["signals"])
            self.assertIn("SIGKILL", termination["signals"])

    def test_unavailable_file_profiler_removes_only_empty_owned_scratch(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            destination = root / "output"
            scratch = root / "scratch"
            destination.mkdir()
            scratch.mkdir()
            binary = {"path": "/tmp/docx", "bytes": 1, "sha256": "a" * 64, "executable": True}
            build = {"path": "route-attempts/build-normal.json", "sha256": "b" * 64}
            tool_info = {"name": "perf-stat", "path": "/usr/bin/perf", "status": "available"}

            def refused(command, *, stdout_path, stderr_path, timeout_seconds):
                stdout_path.write_bytes(b"")
                stderr_path.write_text("No permission to enable cycles event.\n", encoding="utf-8")
                return {"returncode": 1, "timed_out": False}

            with patch.object(profile_routes, "_run_command", side_effect=refused):
                result = profile_routes._run_profile(
                    attempt="profile",
                    binary=binary,
                    protocol={},
                    protocol_sha256="c" * 64,
                    build=build,
                    tool="perf-stat",
                    tool_info=tool_info,
                    case_label="s64-a64-short-c64",
                    route_name="file_store",
                    destination=destination,
                    scratch=scratch,
                    timeout_seconds=1,
                )
            self.assertEqual(result["status"], "unavailable")
            self.assertFalse(any(scratch.iterdir()))

    def test_unlaunchable_profiler_is_unavailable_and_cleans_empty_file_scratch(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            destination = root / "output"
            scratch = root / "scratch"
            destination.mkdir()
            scratch.mkdir()
            binary = {"path": "/tmp/docx", "bytes": 1, "sha256": "a" * 64, "executable": True}
            build = {"path": "route-attempts/build-normal.json", "sha256": "b" * 64}
            tool_info = {"name": "strace", "path": "/usr/bin/strace", "status": "available"}

            def missing(command, *, stdout_path, stderr_path, timeout_seconds):
                return {
                    "returncode": None,
                    "timed_out": False,
                    "route_started": False,
                    "launch_error_type": "FileNotFoundError",
                    "launch_error": "FileNotFoundError: strace",
                }

            with patch.object(profile_routes, "_run_command", side_effect=missing):
                result = profile_routes._run_profile(
                    attempt="profile",
                    binary=binary,
                    protocol={},
                    protocol_sha256="c" * 64,
                    build=build,
                    tool="strace",
                    tool_info=tool_info,
                    case_label="s64-a64-short-c64",
                    route_name="file_store",
                    destination=destination,
                    scratch=scratch,
                    timeout_seconds=1,
                )
            self.assertEqual(result["status"], "unavailable")
            self.assertFalse(any(scratch.iterdir()))
            receipt = profile_routes.read(destination / "strace-file_store-s64-a64-short-c64" / "receipt.json")
            self.assertEqual(receipt["failure_class"], "profiler_unavailable")
            self.assertEqual(receipt["process"]["launch_error_type"], "FileNotFoundError")

    def test_heaptrack_print_missing_outputs_leave_counters_unavailable(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            destination = root / "output"
            scratch = root / "scratch"
            destination.mkdir()
            scratch.mkdir()
            binary = {"path": "/tmp/docx", "bytes": 1, "sha256": "a" * 64, "executable": True}
            build = {"path": "route-attempts/build-normal.json", "sha256": "b" * 64}
            tool_info = {"name": "heaptrack", "path": "/usr/bin/heaptrack", "status": "available"}

            def successful_capture(command, *, stdout_path, stderr_path, timeout_seconds):
                if stdout_path.name == "stdout.txt":
                    stdout_path.write_bytes(b"route stdout")
                    stderr_path.write_bytes(b"")
                    stdout_path.parent.joinpath("report.json").write_text("{}", encoding="utf-8")
                    stdout_path.parent.joinpath("resource.txt").write_text("Maximum resident set size (kbytes): 1\n", encoding="utf-8")
                    stdout_path.parent.joinpath("heaptrack-profile.gz").write_bytes(b"capture")
                return {"returncode": 0, "timed_out": False, "route_started": True}

            with (
                patch.object(profile_routes, "_run_command", side_effect=successful_capture),
                patch.object(
                    profile_routes,
                    "_heaptrack_print_info",
                    return_value={"name": "heaptrack_print", "path": "/usr/bin/heaptrack_print", "status": "available"},
                ),
                patch.object(profile_routes.routes, "check_route_report", return_value={}),
            ):
                result = profile_routes._run_profile(
                    attempt="profile",
                    binary=binary,
                    protocol={},
                    protocol_sha256="c" * 64,
                    build=build,
                    tool="heaptrack",
                    tool_info=tool_info,
                    case_label="s64-a64-short-c64",
                    route_name="deterministic",
                    destination=destination,
                    scratch=scratch,
                    timeout_seconds=1,
                )
            self.assertEqual(result["status"], "ok")
            receipt = profile_routes.read(destination / "heaptrack-deterministic-s64-a64-short-c64" / "receipt.json")
            self.assertEqual(receipt["profiler"]["counter_status"], "unavailable")
            self.assertEqual(receipt["profiler"]["summary_status"], "unavailable")
            self.assertEqual(receipt["profiler"]["heaptrack_print_run"]["returncode"], 0)

    def test_file_replay_cleanup_requires_empty_owned_directory(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            run_scratch = root / "run"
            replay = run_scratch / "replay"
            replay.mkdir(parents=True)
            self.assertEqual(profile_routes._cleanup_file_replay(replay, run_scratch), (True, None))
            self.assertFalse(run_scratch.exists())

            run_scratch = root / "retained-run"
            replay = run_scratch / "replay"
            replay.mkdir(parents=True)
            (replay / "unexpected").write_bytes(b"retained")
            cleaned, error = profile_routes._cleanup_file_replay(replay, run_scratch)
            self.assertFalse(cleaned)
            self.assertIn("retained files", error or "")
            self.assertTrue((replay / "unexpected").is_file())

    def test_profile_inventory_defaults_to_required_heavy_cases(self):
        self.assertEqual(profile_routes.PROFILE_CASES, (
            "s131072-a64-short-c64",
            "s64-a16384-short-c64",
        ))
        self.assertEqual(profile_routes.PROFILE_TOOLS, ("perf-stat", "perf-record", "strace", "heaptrack"))
        self.assertEqual(tuple(route.name for route in routes.ROUTES), (
            "deterministic", "memory_store", "file_store",
        ))


if __name__ == "__main__":
    unittest.main()
