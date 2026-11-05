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

    def _analyzer_config(self, root: Path) -> tuple[dict, Path]:
        config = self._config(root)
        output = config["paths"]["output"]
        analyzer = root / "fake-analyzer.py"
        analyzer.write_text(
            "from pathlib import Path\n"
            "import sys\n"
            "Path(sys.argv[1], 'analysis.json').write_text('{\"status\":\"ok\"}')\n",
            encoding="utf-8",
        )
        manifest = {
            "schema": "litchi_file_qualification.v3",
            "status": "complete; analyze.py review pending",
            "routes": [f"route-{index}" for index in range(wrapper.TARGET_ROUTE_COUNT)],
            "calibration": [{} for _ in range(wrapper.TARGET_ROUTE_COUNT)],
            "abba": [{} for _ in range(wrapper.TARGET_ROUTE_COUNT)],
        }
        (output / "manifest.json").write_bytes(wrapper.canonical(manifest))
        for index in range(2 + 2 * wrapper.TARGET_ROUTE_COUNT + 4 * wrapper.ROUNDS * wrapper.TARGET_ROUTE_COUNT):
            (output / f"sample-{index}.receipt.json").write_text("{}", encoding="utf-8")
        config["paths"]["analyzer"] = analyzer
        config["artifacts"] = {"python": {"path": sys.executable}}
        return config, analyzer

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

    def test_invoke_analyzer_happy_path_checks_real_manifest_and_exit_code(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            config, _analyzer = self._analyzer_config(Path(directory))
            state: dict = {}
            owned: dict = {"analyzer": None}
            with mock.patch.object(wrapper, "POLL_SECONDS", 0.01):
                wrapper.invoke_analyzer(config, state, owned)
            self.assertEqual(state["analyzer_exit_code"], 0)
            self.assertIsNotNone(owned["analyzer"])
            self.assertTrue((config["paths"]["output"] / "analysis.json").is_file())
            cleanup = wrapper.cleanup_owned(None, None, None, owned["analyzer"], {"seen_scope_processes": {}, "last_cgroup_path": None})
            self.assertEqual(cleanup["errors"], [])

    def test_invoke_analyzer_records_nonzero_exit_and_fails_closed(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            config, analyzer = self._analyzer_config(Path(directory))
            analyzer.write_text("raise SystemExit(7)\n", encoding="utf-8")
            state: dict = {}
            owned: dict = {"analyzer": None}
            with mock.patch.object(wrapper, "POLL_SECONDS", 0.01), self.assertRaises(wrapper.GuardFailure):
                wrapper.invoke_analyzer(config, state, owned)
            self.assertEqual(state["analyzer_exit_code"], 7)
            self.assertIsNotNone(owned["analyzer"])
            cleanup = wrapper.cleanup_owned(None, None, None, owned["analyzer"], {"seen_scope_processes": {}, "last_cgroup_path": None})
            self.assertEqual(cleanup["errors"], [])

    def test_invoke_analyzer_owns_observed_outside_scope_child(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            config, analyzer = self._analyzer_config(Path(directory))
            analyzer.write_text(
                "from pathlib import Path\n"
                "import subprocess\n"
                "import sys\n"
                "import time\n"
                "output = Path(sys.argv[1])\n"
                "child = subprocess.Popen([sys.executable, '-c', 'import time; time.sleep(30)'], start_new_session=True)\n"
                "(output / 'analyzer-child.pid').write_text(str(child.pid))\n"
                "(output / 'analysis.json').write_text('{\"status\":\"ok\"}')\n"
                "time.sleep(0.2)\n",
                encoding="utf-8",
            )
            state: dict = {}
            owned: dict = {"analyzer": None}
            try:
                with mock.patch.object(wrapper, "POLL_SECONDS", 0.01):
                    wrapper.invoke_analyzer(config, state, owned)
                child_pid = int((config["paths"]["output"] / "analyzer-child.pid").read_text())
                self.assertTrue(state.get("seen_analyzer_processes"))
                cleanup = wrapper.cleanup_owned(
                    None,
                    None,
                    None,
                    owned["analyzer"],
                    {"seen_analyzer_processes": state["seen_analyzer_processes"], "last_cgroup_path": None},
                )
                self.assertEqual(cleanup["errors"], [])
                self.assertIsNone(wrapper.proc_identity(child_pid))
            finally:
                if owned.get("analyzer") is not None and owned["analyzer"].poll() is None:
                    wrapper.cleanup_owned(None, None, None, owned["analyzer"], state)

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
        self.assertNotIn("--wait", command)
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

    def test_transient_scope_gc_empty_cgroup_retains_metadata_and_skips_kill(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            calls = root / "systemctl-calls"
            kill_log = root / "systemctl-kill"
            systemctl = root / "fake-systemctl-lifecycle.py"
            systemctl.write_text(
                "#!/usr/bin/env python3\n"
                "import sys\n"
                "from pathlib import Path\n"
                f"calls = Path({str(calls)!r})\n"
                f"kill_log = Path({str(kill_log)!r})\n"
                "if 'kill' in sys.argv:\n"
                "    kill_log.write_text('called')\n"
                "    raise SystemExit(1)\n"
                "count = int(calls.read_text() or '0') if calls.exists() else 0\n"
                "calls.write_text(str(count + 1))\n"
                "if count == 0:\n"
                "    print('ActiveState=active')\n"
                "    print('SubState=running')\n"
                "    print('Result=success')\n"
                "    print('ExecMainStatus=0')\n"
                "    print('MainPID=123')\n"
                "    print('ControlGroup=/w8-test.scope')\n"
                "    print('MemoryCurrent=1000')\n"
                "    print('MemoryPeak=1234')\n"
                "else:\n"
                "    print('ActiveState=inactive')\n"
                "    print('SubState=dead')\n"
                "    print('Result=success')\n"
                "    print('ControlGroup=')\n"
                "    print('MemoryCurrent=[not set]')\n"
                "    print('MemoryPeak=[not set]')\n",
                encoding="utf-8",
            )
            systemctl.chmod(systemctl.stat().st_mode | stat.S_IXUSR)
            errors: list[str] = []
            with mock.patch.object(wrapper, "cgroup_pids", return_value=("/w8-test.scope", [], None)):
                active = wrapper.scope_state(systemctl, "w8-test.scope", {}, [])
                gone = wrapper.scope_state(systemctl, "w8-test.scope", {}, errors)
            self.assertTrue(active["available"])
            self.assertFalse(gone["available"])
            self.assertEqual(gone["monitor_error_kind"], "unit-gone")
            self.assertEqual(gone["ActiveState"], "inactive")
            self.assertEqual(gone["ControlGroup"], "")
            self.assertEqual(len(errors), 1)

            incomplete = root / "fake-systemctl-incomplete-gc.py"
            incomplete.write_text(
                "#!/usr/bin/env python3\n"
                "print('ActiveState=inactive')\n"
                "print('SubState=dead')\n"
                "print('ControlGroup=')\n"
                "print('MemoryCurrent=[not set]')\n"
                "print('MemoryPeak=[not set]')\n",
                encoding="utf-8",
            )
            incomplete.chmod(incomplete.stat().st_mode | stat.S_IXUSR)
            incomplete_value = wrapper.show_unit(incomplete, "w8-test.scope")
            self.assertFalse(incomplete_value["available"])
            self.assertEqual(incomplete_value["monitor_error_kind"], "schema")

            config = self._config(root / "checks")
            state = {
                "runner_exit_code": 0,
                "timeout_exit_code": 0,
                "unit_metadata_missing": False,
                "memory_peak_bytes": 1234,
                "memory_current_max_bytes": 0,
                "leftover_processes": [],
                "leftover_cgroup_pids": [],
                "deadline_exceeded": False,
                "output_limit_exceeded": False,
                "monitor_errors": [],
                "post_monitor_errors": errors,
                "post_monitor_available_samples": 0,
                "unit_gc_expected": True,
                "last_cgroup_path": "/w8-test.scope",
            }
            with mock.patch.object(wrapper, "forbidden_rows", return_value=([], [])):
                self.assertEqual(wrapper.post_process_checks(config, state), [])
            with mock.patch.object(
                wrapper,
                "cgroup_pids",
                return_value=("/w8-test.scope", [], "cgroup.procs is gone: scope was garbage-collected"),
            ):
                cleanup = wrapper.cleanup_owned(
                    systemctl,
                    "w8-test.scope",
                    None,
                    None,
                    {"seen_scope_processes": {}, "last_cgroup_path": "/w8-test.scope", "unit_gc_expected": True},
                )
            self.assertEqual(cleanup["errors"], [])
            self.assertFalse(kill_log.exists())

    def test_actual_user_systemd_scope_gc_shape_without_native_runner(self) -> None:
        systemd_run = Path("/usr/bin/systemd-run")
        systemctl = Path("/usr/bin/systemctl")
        if not systemd_run.is_file() or not systemctl.is_file():
            self.skipTest("systemd user tools are unavailable")
        status = subprocess.run(
            [str(systemctl), "--user", "is-system-running"],
            capture_output=True,
            text=True,
            timeout=5,
            check=False,
        )
        if status.returncode != 0 or status.stdout.strip() not in {"running", "degraded"}:
            self.skipTest(f"user systemd is unavailable: {status.stdout.strip()} {status.stderr.strip()}")
        unit = f"w8-quiet-test-{os.getpid()}-{time.time_ns()}.scope"
        command = [
            str(systemd_run),
            "--user",
            "--scope",
            f"--unit={unit}",
            "--property=MemoryAccounting=yes",
            f"--property=MemoryMax={wrapper.MEMORY_MAX_BYTES}",
            f"--property=MemorySwapMax={wrapper.SWAP_MAX_BYTES}",
            "--property=TasksMax=4096",
            "--property=KillMode=control-group",
            "--",
            "/bin/sh",
            "-c",
            "sleep 0.1",
        ]
        completed = subprocess.run(command, capture_output=True, text=True, timeout=15, check=False)
        self.assertEqual(completed.returncode, 0, completed.stderr)
        observed = wrapper.show_unit(systemctl, unit)
        self.assertFalse(observed["available"])
        self.assertEqual(observed["monitor_error_kind"], "unit-gone")
        self.assertEqual(observed["ActiveState"], "inactive")
        self.assertEqual(observed["SubState"], "dead")
        self.assertEqual(observed["Result"], "success")
        self.assertEqual(observed["ControlGroup"], "")
        self.assertEqual(observed["MemoryCurrent"], "[not set]")
        self.assertEqual(observed["MemoryPeak"], "[not set]")
        self.assertNotIn("MainPID", observed)
        self.assertNotIn("ExecMainStatus", observed)

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

    def test_cleanup_preserves_top_level_popen_return_code(self) -> None:
        process = subprocess.Popen(
            [sys.executable, "-c", "import sys; sys.exit(7)"],
            start_new_session=True,
        )
        deadline = time.monotonic() + 5.0
        while True:
            observed = os.waitid(
                os.P_PID,
                process.pid,
                os.WEXITED | os.WNOHANG | os.WNOWAIT,
            )
            if observed is not None and observed.si_pid == process.pid:
                break
            if time.monotonic() >= deadline:
                process.kill()
                process.wait()
                self.fail("top-level Popen child did not exit within the observation deadline")
            time.sleep(0.01)
        identity = wrapper.proc_identity(process.pid)
        self.assertIsNotNone(identity)
        state = {
            "seen_scope_processes": {f"{identity['pid']}:{identity['starttime']}": identity},
            "last_cgroup_path": None,
        }
        cleanup = wrapper.cleanup_owned(None, None, process, None, state)
        self.assertEqual(cleanup["errors"], [])
        self.assertEqual(process.returncode, 7)
        self.assertEqual(process.wait(timeout=0), 7)

    def test_cleanup_reaps_analyzer_process_group(self) -> None:
        with tempfile.TemporaryDirectory() as _directory:
            process = subprocess.Popen(["sleep", "30"], start_new_session=True)
            cleanup = wrapper.cleanup_owned(None, None, None, process, {"seen_scope_processes": {}, "last_cgroup_path": None})
            self.assertEqual(cleanup["errors"], [])
            self.assertIsNotNone(process.poll())

    def test_cleanup_reaps_analyzer_child_in_owned_process_group(self) -> None:
        code = (
            "import subprocess,time\n"
            "child=subprocess.Popen(['sleep','30'])\n"
            "print(child.pid, flush=True)\n"
            "time.sleep(30)\n"
        )
        analyzer = subprocess.Popen(
            [sys.executable, "-c", code], stdout=subprocess.PIPE, text=True, start_new_session=True
        )
        self.assertIsNotNone(analyzer.stdout)
        child_pid = int(analyzer.stdout.readline().strip())
        cleanup = wrapper.cleanup_owned(None, None, None, analyzer, {"seen_scope_processes": {}, "last_cgroup_path": None})
        analyzer.stdout.close()
        self.assertEqual(cleanup["errors"], [])
        self.assertIsNotNone(analyzer.poll())
        self.assertIsNone(wrapper.proc_identity(child_pid))

    def test_cleanup_kills_separate_session_child_through_owned_cgroup(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            kill_log = root / "systemctl-kill"
            pid_file = root / "owned-pids.json"
            systemctl = root / "fake-systemctl-kill-owned.py"
            systemctl.write_text(
                "#!/usr/bin/env python3\n"
                "import json\n"
                "import os\n"
                "import signal\n"
                "import sys\n"
                "from pathlib import Path\n"
                f"kill_log = Path({str(kill_log)!r})\n"
                f"pid_file = Path({str(pid_file)!r})\n"
                "if 'kill' in sys.argv:\n"
                "    for pid in json.loads(pid_file.read_text()):\n"
                "        try:\n"
                "            os.kill(pid, signal.SIGKILL)\n"
                "        except ProcessLookupError:\n"
                "            pass\n"
                "    kill_log.write_text('called')\n"
                "    raise SystemExit(0)\n"
                "raise SystemExit(0)\n",
                encoding="utf-8",
            )
            systemctl.chmod(systemctl.stat().st_mode | stat.S_IXUSR)
            leader = subprocess.Popen([sys.executable, "-c", "import time; time.sleep(30)"], start_new_session=True)
            child = subprocess.Popen([sys.executable, "-c", "import time; time.sleep(30)"], start_new_session=True)
            self.assertNotEqual(os.getsid(leader.pid), os.getsid(child.pid))
            pid_file.write_text(json.dumps([leader.pid, child.pid]), encoding="utf-8")

            def owned_cgroup(_properties: dict) -> tuple[str, list[int], None]:
                pids = [process.pid for process in (leader, child) if process.poll() is None]
                return "/owned-w8.scope", pids, None

            with mock.patch.object(wrapper, "cgroup_pids", side_effect=owned_cgroup):
                cleanup = wrapper.cleanup_owned(
                    systemctl,
                    "w8-file-abba-test.scope",
                    leader,
                    None,
                    {"seen_scope_processes": {}, "last_cgroup_path": "/owned-w8.scope"},
                )
            self.assertEqual(cleanup["errors"], [])
            self.assertTrue(kill_log.exists())
            self.assertIsNotNone(leader.poll())
            self.assertIsNotNone(child.poll())
            self.assertEqual(cleanup["remaining_cgroup_pids"], [])

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
                {
                    "seen_scope_processes": {},
                    "last_cgroup_path": "/w8-test.scope",
                    "unit_gc_expected": True,
                },
            )
        self.assertTrue(any("permission denied" in error for error in cleanup["errors"]))

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
