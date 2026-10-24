"""Source-only lifecycle tests for the guarded W8 ABBA launcher.

These tests never invoke Cargo, a native probe, a profiler, or systemd.  They
use temporary fake commands and synthetic lifecycle state to exercise the
normal, failed, timed-out, escaped-child, and oversized-output refusals.
"""
from __future__ import annotations

import importlib.util
import json
import os
from pathlib import Path
import stat
import subprocess
import sys
import tempfile
import time
import unittest
from unittest import mock


ROOT = Path(__file__).resolve().parents[1]
SCRIPT = ROOT / "docs/unified-ops/results/0362ul32/w8-quiet-abba.py"
spec = importlib.util.spec_from_file_location("w8_quiet_abba", SCRIPT)
if spec is None or spec.loader is None:
    raise AssertionError(f"cannot load {SCRIPT}")
wrapper = importlib.util.module_from_spec(spec)
spec.loader.exec_module(wrapper)


class QuietAbbaWrapperTests(unittest.TestCase):
    def _config(self, root: Path) -> dict:
        output = root / "output"
        tmp = root / "tmp"
        output.mkdir(parents=True)
        tmp.mkdir(parents=True)
        return {
            "paths": {"output": output, "tmp": tmp},
            "memory_max_bytes": wrapper.MEMORY_MAX_BYTES,
            "aggregate_output_limit_bytes": wrapper.AGGREGATE_OUTPUT_LIMIT_BYTES,
            "file_limit_bytes": wrapper.OUTPUT_FILE_LIMIT_BYTES,
        }

    def test_launcher_sets_exact_byte_rlimit_and_no_argument_boundary(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            paths = {
                "runner": root / "run.py",
                "before_binary": root / "before",
                "after_binary": root / "after",
                "output": root / "output",
                "before_manifest": root / "before.json",
                "after_manifest": root / "after.json",
                "policy": root / "policy.json",
            }
            source = wrapper.launcher_source(paths, 7, root / "pair.json", root / "tmp")
            self.assertIn(f"RLIMIT_FSIZE, ({wrapper.OUTPUT_FILE_LIMIT_BYTES}, {wrapper.OUTPUT_FILE_LIMIT_BYTES})", source)
            self.assertIn("accepts no arguments", source)
            self.assertNotIn("ulimit", source)

    def test_launcher_binding_rejects_cpu_or_argv_mutation(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            paths = {
                "runner": root / "run.py",
                "before_binary": root / "before",
                "after_binary": root / "after",
                "output": root / "output",
                "before_manifest": root / "before.json",
                "after_manifest": root / "after.json",
                "policy": root / "policy.json",
            }
            launcher = root / "launcher.py"
            launcher.write_text(wrapper.launcher_source(paths, 7, root / "pair.json", root / "tmp"), encoding="utf-8")
            binding = wrapper.launcher_binding(launcher)
            self.assertEqual(binding["cpu"], 7)
            self.assertEqual(binding["argv"][1], "-B")
            launcher.write_text(launcher.read_text().replace("LAUNCHER_CPU = 7", "LAUNCHER_CPU = 8"), encoding="utf-8")
            self.assertEqual(wrapper.launcher_binding(launcher)["cpu"], 8)
            with self.assertRaises(wrapper.GuardFailure):
                wrapper.verify_launcher_binding(launcher, 7, binding["argv"])
            launcher.write_text(
                wrapper.launcher_source(paths, 7, root / "pair.json", root / "tmp").replace(
                    str(root / "pair.json"), str(root / "other-pair.json")
                ),
                encoding="utf-8",
            )
            with self.assertRaises(wrapper.GuardFailure):
                wrapper.verify_launcher_binding(launcher, 7, binding["argv"])
            launcher.write_text(
                launcher.read_text() + "\nLAUNCHER_CPU = 7\n",
                encoding="utf-8",
            )
            with self.assertRaises(wrapper.GuardFailure):
                wrapper.launcher_binding(launcher)

    def test_fake_launcher_executes_normal_and_rejects_extra_argument(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            runner = root / "fake-runner.py"
            runner.write_text(
                "#!/usr/bin/env python3\n"
                "from pathlib import Path\n"
                "import os\n"
                "Path(os.environ['FAKE_TARGET']).write_text('ok')\n",
                encoding="utf-8",
            )
            runner.chmod(runner.stat().st_mode | stat.S_IXUSR)
            output = root / "result"
            target = root / "target.txt"
            paths = {
                "runner": runner,
                "before_binary": root / "before",
                "after_binary": root / "after",
                "output": output,
                "before_manifest": root / "before.json",
                "after_manifest": root / "after.json",
                "policy": root / "policy.json",
            }
            launcher = root / "launcher.py"
            launcher.write_text(wrapper.launcher_source(paths, 7, root / "pair.json", root / "tmp"), encoding="utf-8")
            launcher.chmod(launcher.stat().st_mode | stat.S_IXUSR)
            environment = os.environ.copy()
            environment["FAKE_TARGET"] = str(target)
            normal = subprocess.run([str(launcher)], check=False, capture_output=True, text=True, env=environment)
            self.assertEqual(normal.returncode, 0, normal.stderr)
            self.assertEqual(target.read_text(), "ok")
            rejected = subprocess.run([str(launcher), "extra"], check=False, capture_output=True, text=True, env=environment)
            self.assertNotEqual(rejected.returncode, 0)

    def test_systemd_command_has_exact_cgroup_timeout_and_runner_launcher_only(self) -> None:
        config = {
            "unit": "w8-file-abba-test.scope",
            "memory_max_bytes": 20_000_000_000,
            "swap_max_bytes": 0,
            "stop_grace_seconds": 60,
            "deadline_seconds": 43_200,
            "artifacts": {
                "systemd_run": {"path": "/usr/bin/systemd-run"},
                "timeout": {"path": "/usr/bin/timeout"},
            },
        }
        paths = {"launcher": Path("/external/w8-noarg-launcher.py")}
        command = wrapper.systemd_command(config, paths)
        self.assertIn("--property=MemoryMax=20000000000", command)
        self.assertIn("--property=MemorySwapMax=0", command)
        self.assertIn("--property=KillMode=control-group", command)
        self.assertIn("43200s", command)
        self.assertEqual(command[-1], "/external/w8-noarg-launcher.py")
        self.assertNotIn("file_w8_qualified_probe", " ".join(command))

    def test_fake_systemctl_normal_and_failure_are_recorded(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            normal = root / "fake-systemctl-normal.py"
            normal.write_text(
                "#!/usr/bin/env python3\n"
                "print('ActiveState=active')\n"
                "print('SubState=running')\n"
                "print('Result=success')\n"
                "print('ExecMainStatus=0')\n"
                "print('MainPID=123')\n"
                "print('ControlGroup=/w8-test.scope')\n"
                "print('MemoryCurrent=1000')\n"
                "print('MemoryPeak=1234')\n",
                encoding="utf-8",
            )
            normal.chmod(normal.stat().st_mode | stat.S_IXUSR)
            value = wrapper.show_unit(normal, "w8-test.scope")
            self.assertTrue(value["available"])
            self.assertEqual(value["MemoryPeak"], "1234")

            failure = root / "fake-systemctl-failure.py"
            failure.write_text("#!/usr/bin/env python3\nraise SystemExit(5)\n", encoding="utf-8")
            failure.chmod(failure.stat().st_mode | stat.S_IXUSR)
            value = wrapper.show_unit(failure, "w8-test.scope")
            self.assertFalse(value["available"])
            self.assertEqual(value["returncode"], 5)

            gone = root / "fake-systemctl-unit-gone.py"
            gone.write_text(
                "#!/usr/bin/env python3\n"
                "import sys\n"
                "print('Unit w8-test.scope not loaded.', file=sys.stderr)\n"
                "raise SystemExit(1)\n",
                encoding="utf-8",
            )
            gone.chmod(gone.stat().st_mode | stat.S_IXUSR)
            value = wrapper.show_unit(gone, "w8-test.scope")
            self.assertFalse(value["available"])
            self.assertEqual(value["monitor_error_kind"], "unit-gone")

            missing = root / "fake-systemctl-missing-control-group.py"
            missing.write_text(
                "#!/usr/bin/env python3\n"
                "print('ActiveState=active')\n"
                "print('SubState=running')\n"
                "print('Result=success')\n"
                "print('ExecMainStatus=0')\n"
                "print('MainPID=123')\n"
                "print('MemoryCurrent=1000')\n"
                "print('MemoryPeak=1234')\n",
                encoding="utf-8",
            )
            missing.chmod(missing.stat().st_mode | stat.S_IXUSR)
            value = wrapper.show_unit(missing, "w8-test.scope")
            self.assertFalse(value["available"])
            self.assertEqual(value["monitor_error_kind"], "schema")
            _path, _pids, error = wrapper.cgroup_pids({"ControlGroup": "/missing-w8-cgroup"})
            self.assertIsNotNone(error)

    def test_normal_state_passes_output_and_descendant_guards(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            config = self._config(Path(directory))
            state = {
                "runner_exit_code": 0,
                "timeout_exit_code": 0,
                "unit_metadata_missing": False,
                "memory_peak_bytes": 1024,
                "memory_current_max_bytes": 1024,
                "leftover_processes": [],
                "leftover_cgroup_pids": [],
                "deadline_exceeded": False,
                "output_limit_exceeded": False,
            }
            with mock.patch.object(wrapper, "forbidden_rows", return_value=([], [])):
                self.assertEqual(wrapper.post_process_checks(config, state), [])

    def test_active_monitor_error_fails_after_a_good_sample(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            config = self._config(Path(directory))
            state = {
                "runner_exit_code": 0,
                "timeout_exit_code": 0,
                "unit_metadata_missing": False,
                "memory_peak_bytes": 1024,
                "memory_current_max_bytes": 1024,
                "leftover_processes": [],
                "leftover_cgroup_pids": [],
                "deadline_exceeded": False,
                "output_limit_exceeded": False,
                "monitor_errors": ["systemctl: intermittent failure"],
            }
            with mock.patch.object(wrapper, "forbidden_rows", return_value=([], [])):
                failures = wrapper.post_process_checks(config, state)
            self.assertTrue(any("active unit monitor errors" in failure for failure in failures))

    def test_scope_state_retains_missing_and_cgroup_read_errors(self) -> None:
        errors: list[str] = []
        with mock.patch.object(
            wrapper,
            "show_unit",
            return_value={
                "available": False,
                "monitor_error_kind": "schema",
                "monitor_error": "systemctl omitted ControlGroup",
            },
        ):
            value = wrapper.scope_state(Path("/systemctl"), "w8-test.scope", {}, errors)
        self.assertFalse(value["available"])
        self.assertEqual(errors, ["schema: systemctl omitted ControlGroup"])

        errors = []
        with mock.patch.object(
            wrapper,
            "show_unit",
            return_value={"available": True, "ControlGroup": "/w8-test.scope"},
        ), mock.patch.object(
            wrapper,
            "cgroup_pids",
            return_value=("/w8-test.scope", [], "cannot read cgroup.procs: permission denied"),
        ):
            value = wrapper.scope_state(Path("/systemctl"), "w8-test.scope", {}, errors)
        self.assertFalse(value["available"])
        self.assertEqual(errors, ["cgroup: cannot read cgroup.procs: permission denied"])

    def test_failed_and_timeout_states_fail_closed(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            for field, value in (("runner_exit_code", 7), ("timeout_exit_code", 124), ("deadline_exceeded", True)):
                config = self._config(Path(directory) / field)
                state = {
                    "runner_exit_code": 0,
                    "timeout_exit_code": 0,
                    "unit_metadata_missing": False,
                    "memory_peak_bytes": 1024,
                    "memory_current_max_bytes": 1024,
                    "leftover_processes": [],
                    "leftover_cgroup_pids": [],
                    "deadline_exceeded": False,
                    "output_limit_exceeded": False,
                }
                state[field] = value
                with self.subTest(field=field), mock.patch.object(wrapper, "forbidden_rows", return_value=([], [])):
                    self.assertTrue(wrapper.post_process_checks(config, state))

    def test_escaped_child_and_oversized_output_fail_closed(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            config = self._config(Path(directory))
            escaped = {
                "runner_exit_code": 0,
                "timeout_exit_code": 0,
                "unit_metadata_missing": False,
                "memory_peak_bytes": 1024,
                "memory_current_max_bytes": 1024,
                "leftover_processes": [{"pid": 123, "starttime": "1", "cmdline": "escaped"}],
                "leftover_cgroup_pids": [123],
                "deadline_exceeded": False,
                "output_limit_exceeded": False,
            }
            with mock.patch.object(wrapper, "forbidden_rows", return_value=([], [])):
                self.assertTrue(wrapper.post_process_checks(config, escaped))

            # A sparse file makes the guard test cheap while still exceeding
            # the exact byte-valued per-file limit.
            output_file = config["paths"]["output"] / "analysis.json"
            with output_file.open("wb") as stream:
                stream.truncate(wrapper.OUTPUT_FILE_LIMIT_BYTES + 1)
            normal = {
                "runner_exit_code": 0,
                "timeout_exit_code": 0,
                "unit_metadata_missing": False,
                "memory_peak_bytes": 1024,
                "leftover_processes": [],
                "leftover_cgroup_pids": [],
                "deadline_exceeded": False,
                "output_limit_exceeded": False,
            }
            with mock.patch.object(wrapper, "forbidden_rows", return_value=([], [])):
                failures = wrapper.post_process_checks(config, normal)
            self.assertTrue(any("RLIMIT_FSIZE" in failure for failure in failures))

            tmp_file = config["paths"]["tmp"] / "analysis.json"
            output_file.unlink()
            with tmp_file.open("wb") as stream:
                stream.truncate(wrapper.OUTPUT_FILE_LIMIT_BYTES + 1)
            with mock.patch.object(wrapper, "forbidden_rows", return_value=([], [])):
                failures = wrapper.post_process_checks(config, normal)
            self.assertTrue(any("RLIMIT_FSIZE" in failure for failure in failures))

            tmp_file.unlink()
            with tmp_file.open("wb") as stream:
                stream.truncate(wrapper.AGGREGATE_OUTPUT_LIMIT_BYTES + 1)
            with mock.patch.object(wrapper, "forbidden_rows", return_value=([], [])):
                failures = wrapper.post_process_checks(config, normal)
            self.assertTrue(any("aggregate output exceeds cap" in failure for failure in failures))

    def test_post_monitor_error_without_explicit_gc_proof_fails(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            config = self._config(Path(directory))
            state = {
                "runner_exit_code": 0,
                "timeout_exit_code": 0,
                "unit_metadata_missing": False,
                "memory_peak_bytes": 1024,
                "memory_current_max_bytes": 1024,
                "leftover_processes": [],
                "leftover_cgroup_pids": [],
                "deadline_exceeded": False,
                "output_limit_exceeded": False,
                "post_monitor_errors": ["systemctl: temporary failure"],
                "post_monitor_available_samples": 0,
                "unit_gc_expected": False,
            }
            with mock.patch.object(wrapper, "forbidden_rows", return_value=([], [])):
                failures = wrapper.post_process_checks(config, state)
            self.assertTrue(any("post-run unit monitor errors" in failure for failure in failures))

    def test_only_explicit_unit_disappearance_after_valid_sample_is_tolerated(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            config = self._config(Path(directory))
            state = {
                "runner_exit_code": 0,
                "timeout_exit_code": 0,
                "unit_metadata_missing": False,
                "memory_peak_bytes": 1024,
                "memory_current_max_bytes": 1024,
                "leftover_processes": [],
                "leftover_cgroup_pids": [],
                "deadline_exceeded": False,
                "output_limit_exceeded": False,
                "post_monitor_errors": ["unit-gone: systemctl failed"],
                "post_monitor_available_samples": 0,
                "unit_gc_expected": True,
            }
            with mock.patch.object(wrapper, "forbidden_rows", return_value=([], [])):
                self.assertEqual(wrapper.post_process_checks(config, state), [])

            state["unit_gc_expected"] = False
            with mock.patch.object(wrapper, "forbidden_rows", return_value=([], [])):
                self.assertTrue(wrapper.post_process_checks(config, state))

    def test_cleanup_reaps_owned_process_group(self) -> None:
        with tempfile.TemporaryDirectory() as _directory:
            process = subprocess.Popen(["sleep", "30"], start_new_session=True)
            identity = wrapper.proc_identity(process.pid)
            self.assertIsNotNone(identity)
            state = {
                "seen_scope_processes": {f"{identity['pid']}:{identity['starttime']}": identity},
                "last_cgroup_path": None,
            }
            cleanup = wrapper.cleanup_owned(None, None, process, None, state)
            self.assertEqual(cleanup["errors"], [])
            self.assertIsNotNone(process.poll())
            self.assertEqual(cleanup["remaining_processes"], [])

    def test_cleanup_reaps_analyzer_process_group(self) -> None:
        with tempfile.TemporaryDirectory() as _directory:
            process = subprocess.Popen(["sleep", "30"], start_new_session=True)
            cleanup = wrapper.cleanup_owned(None, None, None, process, {"seen_scope_processes": {}, "last_cgroup_path": None})
            self.assertEqual(cleanup["errors"], [])
            self.assertIsNotNone(process.poll())

    def test_cleanup_fails_closed_on_unreadable_owned_cgroup(self) -> None:
        with mock.patch.object(
            wrapper,
            "cgroup_pids",
            return_value=("/w8-test.scope", [], "cannot read cgroup.procs: permission denied"),
        ):
            cleanup = wrapper.cleanup_owned(
                None,
                None,
                None,
                None,
                {"seen_scope_processes": {}, "last_cgroup_path": "/w8-test.scope"},
            )
        self.assertTrue(any("cannot read cgroup.procs" in error for error in cleanup["errors"]))

    def test_cleanup_reaps_escaped_child_after_leader_exit(self) -> None:
        code = (
            "import subprocess,sys,time\n"
            "child=subprocess.Popen(['sleep','30'])\n"
            "print(child.pid, flush=True)\n"
            "time.sleep(0.05)\n"
        )
        leader = subprocess.Popen(
            [sys.executable, "-c", code],
            stdout=subprocess.PIPE,
            text=True,
            start_new_session=True,
        )
        self.assertIsNotNone(leader.stdout)
        escaped_pid = int(leader.stdout.readline().strip())
        self.assertEqual(leader.wait(timeout=5), 0)
        leader.stdout.close()
        cleanup = wrapper.cleanup_owned(None, None, leader, None, {"seen_scope_processes": {}, "last_cgroup_path": None})
        self.assertEqual(cleanup["errors"], [])
        for _ in range(20):
            if wrapper.proc_identity(escaped_pid) is None:
                break
            time.sleep(0.05)
        self.assertIsNone(wrapper.proc_identity(escaped_pid))


if __name__ == "__main__":
    unittest.main()
