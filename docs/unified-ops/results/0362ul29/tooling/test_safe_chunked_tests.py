from __future__ import annotations

import hashlib
import importlib.util
import json
import os
from pathlib import Path
import stat
import subprocess
import sys
import tempfile
import textwrap
import threading
import time
import unittest
from unittest import mock


MODULE_PATH = Path(__file__).with_name("safe_chunked_tests.py")
SPEC = importlib.util.spec_from_file_location("safe_chunked_tests_under_test", MODULE_PATH)
assert SPEC and SPEC.loader
safe = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(safe)

RUNNER_PATH = MODULE_PATH.with_name("run_safe_gate.py")
RUNNER_SPEC = importlib.util.spec_from_file_location("run_safe_gate_under_test", RUNNER_PATH)
assert RUNNER_SPEC and RUNNER_SPEC.loader
runner = importlib.util.module_from_spec(RUNNER_SPEC)
RUNNER_SPEC.loader.exec_module(runner)


FAKE_COMMAND = textwrap.dedent(
    """
    import json
    from pathlib import Path
    import sys
    import time

    artifact = Path(sys.argv[1])
    marker = Path(sys.argv[2])
    artifact.parent.mkdir(parents=True, exist_ok=True)
    artifact.write_bytes(b"\x7fELFfake executable bytes")
    artifact.chmod(0o751)
    sidecar = artifact.with_suffix(".d")
    sidecar.write_bytes(b"fake private dependency sidecar")
    print(json.dumps({
        "reason": "compiler-artifact",
        "target": {"kind": ["example"]},
        "executable": str(artifact),
        "filenames": [str(artifact), str(sidecar)],
    }), flush=True)
    time.sleep(0.15)
    marker.write_text(str(artifact.exists()), encoding="utf-8")
    """
)


class SafeChunkedLifecycleTests(unittest.TestCase):
    def setUp(self) -> None:
        self.values = {
            "MAX_TARGET": safe.MAX_TARGET,
            "LIVE_TARGET": safe.LIVE_TARGET,
            "START_TARGET": safe.START_TARGET,
            "MIN_FREE": safe.MIN_FREE,
            "DESCENDANT_WAIT_SECONDS": safe.DESCENDANT_WAIT_SECONDS,
            "PROCESS_KILL_GRACE_SECONDS": safe.PROCESS_KILL_GRACE_SECONDS,
        }

    def tearDown(self) -> None:
        for name, value in self.values.items():
            setattr(safe, name, value)

    def _command(self, directory: Path, script: str) -> tuple[list[str], Path, Path, Path, Path, Path]:
        script_path = directory / "fake_command.py"
        script_path.write_text(script, encoding="utf-8")
        target = directory / "target"
        artifact = target / "debug" / "examples" / "fake-example"
        marker = directory / "marker"
        quarantine = directory / "quarantine"
        log = directory / "command.log"
        return [sys.executable, str(script_path), str(artifact), str(marker)], target, artifact, marker, quarantine, log

    def test_retirement_waits_for_child_and_preserves_hash_mode_and_manifest(self) -> None:
        with tempfile.TemporaryDirectory(prefix="safe-chunked-lifecycle-") as raw:
            command, target, artifact, marker, quarantine, log = self._command(Path(raw), FAKE_COMMAND)
            safe.MIN_FREE = 0
            status, retired, _ = safe.execute(command, target, log, quarantine, 0)
            self.assertEqual(status, 0)
            self.assertEqual(marker.read_text(encoding="utf-8"), "True")
            self.assertEqual(len(retired), 2)
            record = retired[0]
            destination = Path(record["quarantine"])
            self.assertFalse(artifact.exists())
            self.assertFalse(artifact.with_suffix(".d").exists())
            self.assertEqual(destination.read_bytes(), b"\x7fELFfake executable bytes")
            self.assertEqual(record["sha256"], hashlib.sha256(destination.read_bytes()).hexdigest())
            self.assertEqual(record["mode"], "0751")
            self.assertEqual(stat.S_IMODE(destination.stat().st_mode), 0o751)
            manifest = quarantine / "manifest.jsonl"
            manifest_rows = [json.loads(line) for line in manifest.read_text(encoding="utf-8").splitlines()]
            self.assertEqual(manifest_rows[0]["quarantine"], str(destination))

    def test_shared_quarantine_namespaces_same_chunk_and_artifact_by_suite(self) -> None:
        with tempfile.TemporaryDirectory(prefix="safe-chunked-shared-quarantine-") as raw:
            directory = Path(raw)
            command, target, artifact, marker, quarantine, log = self._command(directory, FAKE_COMMAND)
            safe.MIN_FREE = 0
            first_status, first_retired, _ = safe.execute(
                command,
                target,
                log,
                quarantine,
                0,
                quarantine_scope="suite-a",
            )
            self.assertEqual(first_status, 0)
            self.assertEqual(len(first_retired), 2)
            second_status, second_retired, _ = safe.execute(
                command,
                target,
                log,
                quarantine,
                0,
                quarantine_scope="suite-b",
            )
            self.assertEqual(second_status, 0)
            self.assertEqual(len(second_retired), 2)
            first_paths = {row["quarantine"] for row in first_retired}
            second_paths = {row["quarantine"] for row in second_retired}
            self.assertTrue(first_paths.isdisjoint(second_paths))
            self.assertEqual({row["quarantine_scope"] for row in first_retired}, {"suite-a"})
            self.assertEqual({row["quarantine_scope"] for row in second_retired}, {"suite-b"})
            rows = [json.loads(line) for line in (quarantine / "manifest.jsonl").read_text(encoding="utf-8").splitlines()]
            self.assertEqual(len(rows), 4)
            self.assertEqual({row["quarantine_namespace"] for row in rows}, {safe._quarantine_namespace("suite-a"), safe._quarantine_namespace("suite-b")})

    def test_existing_destination_rejects_corruption_without_overwrite(self) -> None:
        with tempfile.TemporaryDirectory(prefix="safe-chunked-corrupt-destination-") as raw:
            directory = Path(raw)
            command, target, artifact, marker, quarantine, log = self._command(directory, FAKE_COMMAND)
            safe.MIN_FREE = 0
            status, retired, _ = safe.execute(command, target, log, quarantine, 0, quarantine_scope="suite")
            self.assertEqual(status, 0)
            self.assertEqual(len(retired), 2)
            destination = Path(retired[0]["quarantine"])
            destination.write_bytes(b"corrupt destination")
            with self.assertRaisesRegex(RuntimeError, "differs from source artifact"):
                safe.execute(command, target, log, quarantine, 0, quarantine_scope="suite")
            self.assertTrue(artifact.exists())
            rows = (quarantine / "manifest.jsonl").read_text(encoding="utf-8").splitlines()
            self.assertEqual(len(rows), 2)

    def test_nonzero_child_keeps_artifact_in_target(self) -> None:
        script = FAKE_COMMAND + "\nsys.exit(7)\n"
        with tempfile.TemporaryDirectory(prefix="safe-chunked-failure-") as raw:
            command, target, artifact, marker, quarantine, log = self._command(Path(raw), script)
            safe.MIN_FREE = 0
            status, retired, _ = safe.execute(command, target, log, quarantine, 0)
            self.assertEqual(status, 7)
            self.assertEqual(retired, [])
            self.assertTrue(artifact.exists())
            self.assertFalse((quarantine / "manifest.jsonl").exists())

    def test_descendant_finishes_before_retirement(self) -> None:
        script = textwrap.dedent(
            """
            import json
            from pathlib import Path
            import subprocess
            import sys

            artifact = Path(sys.argv[1])
            marker = Path(sys.argv[2])
            artifact.parent.mkdir(parents=True, exist_ok=True)
            artifact.write_bytes(b"\x7fELFdescendant executable")
            artifact.chmod(0o751)
            subprocess.Popen(
                [sys.executable, "-c", "import os,pathlib,sys,time; os.setsid(); time.sleep(.35); pathlib.Path(sys.argv[1]).write_text('done')", str(marker)],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
            )
            print(json.dumps({"reason": "compiler-artifact", "target": {"kind": ["test"]}, "executable": str(artifact)}), flush=True)
            """
        )
        with tempfile.TemporaryDirectory(prefix="safe-chunked-descendant-") as raw:
            command, target, artifact, marker, quarantine, log = self._command(Path(raw), script)
            safe.MIN_FREE = 0
            status, retired, _ = safe.execute(command, target, log, quarantine, 0)
            self.assertEqual(status, 0)
            self.assertEqual(marker.read_text(encoding="utf-8"), "done")
            self.assertEqual(len(retired), 1)
            self.assertFalse(artifact.exists())

    def test_hung_setsid_descendant_forces_blocked_retention(self) -> None:
        script = textwrap.dedent(
            """
            import json
            import os
            from pathlib import Path
            import subprocess
            import sys
            import time

            artifact = Path(sys.argv[1])
            marker = Path(sys.argv[2])
            artifact.parent.mkdir(parents=True, exist_ok=True)
            artifact.write_bytes(b"\x7fELFhung descendant artifact")
            artifact.chmod(0o751)
            subprocess.Popen(
                [sys.executable, "-c", "import os,pathlib,sys,time; os.setsid(); time.sleep(5); pathlib.Path(sys.argv[1]).write_text('late')", str(marker)],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
            )
            print(json.dumps({"reason": "compiler-artifact", "target": {"kind": ["test"]}, "executable": str(artifact)}), flush=True)
            """
        )
        with tempfile.TemporaryDirectory(prefix="safe-chunked-hung-descendant-") as raw:
            command, target, artifact, marker, quarantine, log = self._command(Path(raw), script)
            safe.MIN_FREE = 0
            safe.DESCENDANT_WAIT_SECONDS = 0.15
            safe.PROCESS_KILL_GRACE_SECONDS = 0.1
            with self.assertRaisesRegex(RuntimeError, "forced cleanup"):
                safe.execute(command, target, log, quarantine, 0)
            self.assertTrue(artifact.exists())
            self.assertFalse(marker.exists())
            self.assertFalse((quarantine / "manifest.jsonl").exists())
            self.assertIn('"forced_cleanup": true', log.read_text(encoding="utf-8"))

    def test_live_overflow_stops_command_and_does_not_skip_or_retire_work(self) -> None:
        script = textwrap.dedent(
            """
            import json
            from pathlib import Path
            import sys
            import time

            artifact = Path(sys.argv[1])
            marker = Path(sys.argv[2])
            artifact.parent.mkdir(parents=True, exist_ok=True)
            artifact.write_bytes(b"\x7fELFnot retired")
            artifact.chmod(0o751)
            print(json.dumps({"reason": "compiler-artifact", "target": {"kind": ["test"]}, "executable": str(artifact)}), flush=True)
            growth = artifact.parents[2] / "growth.bin"
            for _ in range(100):
                with growth.open("ab") as stream:
                    stream.write(b"x" * 8192)
                time.sleep(0.05)
            marker.write_text("after", encoding="utf-8")
            """
        )
        with tempfile.TemporaryDirectory(prefix="safe-chunked-overflow-") as raw:
            directory = Path(raw)
            command, target, artifact, marker, quarantine, log = self._command(directory, script)
            initial = safe.target_bytes(target)
            safe.MIN_FREE = 0
            safe.START_TARGET = initial + 1_000_000
            safe.LIVE_TARGET = initial + 32_000
            safe.MAX_TARGET = initial + 2_000_000
            with self.assertRaises(safe.TargetOverflow):
                safe.execute(command, target, log, quarantine, 0)
            self.assertFalse(marker.exists())
            self.assertTrue(artifact.exists())
            self.assertFalse((quarantine / "manifest.jsonl").exists())
            self.assertIn("safe_chunked_live_guard", log.read_text(encoding="utf-8"))

    def test_plan_partitions_every_test_and_example_without_skipping(self) -> None:
        metadata = {
            "packages": [
                {
                    "name": "fixture",
                    "targets": [
                        {"kind": ["lib"], "test": True, "doctest": True},
                        *({"kind": ["test"], "name": f"test_{i}", "test": True} for i in range(5)),
                        *({"kind": ["example"], "name": f"example_{i}", "test": False} for i in range(4)),
                    ],
                }
            ]
        }
        commands = safe.plan(metadata, {"fixture"}, [], 2)
        rendered = " ".join(" ".join(arguments) for _, arguments in commands)
        for name in [*(f"test_{i}" for i in range(5)), *(f"example_{i}" for i in range(4))]:
            self.assertEqual(rendered.count(name), 1)
        self.assertEqual(sum(1 for verb, _ in commands if verb == "test"), 5)
        self.assertEqual(sum(1 for verb, _ in commands if verb == "build"), 2)

    def test_symlinked_artifact_is_not_moved(self) -> None:
        with tempfile.TemporaryDirectory(prefix="safe-chunked-symlink-") as raw:
            directory = Path(raw)
            target = directory / "target"
            real = target / "debug" / "deps" / "real"
            real.parent.mkdir(parents=True)
            real.write_bytes(b"real")
            real.chmod(0o751)
            link = target / "debug" / "deps" / "link"
            link.symlink_to(real)
            quarantine = directory / "quarantine"
            manifest = directory / "manifest.jsonl"
            safe._retire_artifacts({link}, target, quarantine, manifest, ["fake"], 0, scan_private=False)
            self.assertTrue(link.is_symlink())
            self.assertFalse(manifest.exists())

    def test_retirement_accepts_only_private_elfs_and_exact_sidecars(self) -> None:
        with tempfile.TemporaryDirectory(prefix="safe-chunked-private-outputs-") as raw:
            directory = Path(raw)
            target = directory / "target"
            debug = target / "debug"
            build_output = debug / "build" / "fake-build-script"
            non_elf = debug / "examples" / "fake-example"
            real = debug / "examples" / "real-example"
            shared = debug / "deps" / "libshared.so"
            for path in (build_output, non_elf, real, shared):
                path.parent.mkdir(parents=True, exist_ok=True)
                path.write_bytes(b"\x7fELFreal" if path in {real, shared} else b"not an ELF")
                path.chmod(0o751)
            sidecar = real.with_suffix(".d")
            sidecar.write_text("real sidecar", encoding="utf-8")
            quarantine = directory / "quarantine"
            manifest = directory / "manifest.jsonl"
            retired = safe._retire_artifacts(
                {build_output, non_elf, real},
                target,
                quarantine,
                manifest,
                ["fake"],
                0,
                scan_private=True,
            )
            self.assertEqual({Path(row["relative_source"]).name for row in retired}, {"real-example", "real-example.d"})
            self.assertTrue(build_output.exists())
            self.assertTrue(non_elf.exists())
            self.assertTrue(shared.exists())
            self.assertFalse(real.exists())
            self.assertFalse(sidecar.exists())

    def test_main_stops_after_first_failed_chunk(self) -> None:
        metadata = {
            "workspace_members": ["fixture-id"],
            "target_directory": "/tmp/safe-chunked-fake-target",
            "packages": [{"id": "fixture-id", "name": "fixture", "targets": []}],
        }
        with tempfile.TemporaryDirectory(prefix="safe-chunked-fail-stop-") as raw:
            directory = Path(raw)
            target = directory / "target"
            log_dir = directory / "logs"
            quarantine = directory / "quarantine"
            commands = [("test", ["first"]), ("test", ["second"])]

            def fake_execute(command, _target, log, _quarantine, index, **_kwargs):
                log.write_text(f"chunk {index}\n", encoding="utf-8")
                return (7 if index == 0 else 0), [], {}

            argv = [
                "safe_chunked_tests.py",
                "-p",
                "fixture",
                "--log-dir",
                str(log_dir),
                "--quarantine-dir",
                str(quarantine),
            ]
            with (
                mock.patch.object(safe.sys, "argv", argv),
                mock.patch.object(safe.subprocess, "check_output", return_value=json.dumps(metadata)),
                mock.patch.object(safe, "plan", return_value=commands),
                mock.patch.object(safe, "execute", side_effect=fake_execute) as execute,
            ):
                self.assertEqual(safe.main(), 1)
            self.assertEqual(execute.call_count, 1)
            rows = json.loads((log_dir / "checks.json").read_text(encoding="utf-8"))
            self.assertEqual(len(rows), 1)
            self.assertEqual(rows[0]["status"], "failed")

    def test_helper_unexpected_guard_exception_blocks_and_retains_artifact(self) -> None:
        with tempfile.TemporaryDirectory(prefix="safe-chunked-guard-exception-") as raw:
            directory = Path(raw)
            command, target, artifact, marker, quarantine, log = self._command(
                directory,
                "import time; time.sleep(5)\n",
            )
            artifact.parent.mkdir(parents=True, exist_ok=True)
            artifact.write_bytes(b"\x7fELFretained after guard failure")
            artifact.chmod(0o751)
            safe.MIN_FREE = 0
            with mock.patch.object(safe, "_watch_target", side_effect=RuntimeError("injected helper guard failure")):
                with self.assertRaisesRegex(RuntimeError, "target guard failed"):
                    safe.execute(command, target, log, quarantine, 0)
            self.assertTrue(artifact.exists())
            self.assertFalse((quarantine / "manifest.jsonl").exists())
            self.assertIn("safe_chunked_guard_failure", log.read_text(encoding="utf-8"))

    def test_gate_runner_live_guard_stops_a_non_chunk_command(self) -> None:
        with tempfile.TemporaryDirectory(prefix="safe-gate-live-guard-") as raw:
            directory = Path(raw)
            target = directory / "target"
            script = directory / "grow.py"
            growth = target / "debug" / "growth.bin"
            script.write_text(
                textwrap.dedent(
                    """
                    from pathlib import Path
                    import sys
                    import time

                    path = Path(sys.argv[1])
                    path.parent.mkdir(parents=True, exist_ok=True)
                    while True:
                        with path.open("ab") as stream:
                            stream.write(b"x" * 4096)
                        time.sleep(0.01)
                    """
                ),
                encoding="utf-8",
            )
            initial = runner.target_bytes(target)
            previous_live = runner.LIVE_TARGET
            runner.LIVE_TARGET = initial + 32_000
            process = subprocess.Popen(
                [sys.executable, str(script), str(growth)],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                start_new_session=True,
            )
            stop = threading.Event()
            overflow: list[dict] = []
            stats: dict = {}
            guard = threading.Thread(target=lambda: stats.update(runner._watch_target(process, target, stop, overflow)))
            try:
                guard.start()
                guard.join(timeout=4)
                self.assertFalse(guard.is_alive())
                process.wait(timeout=2)
                self.assertTrue(overflow)
                self.assertGreaterEqual(overflow[0]["target_bytes"], runner.LIVE_TARGET)
            finally:
                stop.set()
                if process.poll() is None:
                    runner._terminate_process_group(process)
                process.wait()
                guard.join(timeout=2)
                runner.LIVE_TARGET = previous_live

    def test_target_measurement_retries_a_transient_du_race(self) -> None:
        with tempfile.TemporaryDirectory(prefix="safe-gate-du-retry-") as raw:
            target = Path(raw) / "target"
            target.mkdir()
            results = [
                subprocess.CompletedProcess(["du"], 1, "", "du: disappearing file\n"),
                subprocess.CompletedProcess(["du"], 0, "123\t%s\n" % target, ""),
            ]
            with mock.patch.object(runner.subprocess, "run", side_effect=results) as run:
                self.assertEqual(runner.target_bytes(target), 123)
            self.assertEqual(run.call_count, 2)

    def test_target_measurement_tolerates_concurrent_create_delete_rename(self) -> None:
        with tempfile.TemporaryDirectory(prefix="safe-gate-du-concurrent-") as raw:
            target = Path(raw) / "target"
            target.mkdir()
            work = target / "debug" / "race"
            work.mkdir(parents=True)
            stop = threading.Event()

            def mutate() -> None:
                for index in range(500):
                    left = work / f"left-{index}"
                    right = work / f"right-{index}"
                    left.write_bytes(b"x" * 128)
                    left.replace(right)
                    right.unlink()
                stop.set()

            worker = threading.Thread(target=mutate)
            worker.start()
            samples = []
            while not stop.is_set():
                samples.append(runner.target_bytes(target))
            worker.join(timeout=2)
            samples.append(runner.target_bytes(target))
            self.assertTrue(samples)
            self.assertTrue(all(value >= 0 for value in samples))

    def test_ownership_excludes_runner_and_ancestors_from_group_matches(self) -> None:
        """A step's process group must not make its coordinator owned."""
        with tempfile.TemporaryDirectory(prefix="safe-gate-ownership-ancestry-") as raw:
            directory = Path(raw)
            target = directory / "target"
            target.mkdir()
            runner_pid = os.getpid()
            ancestor_pid = runner_pid + 10_000_001

            def row(pid: int, ppid: int, pgrp: int, start_time: int, executable: str, cwd: Path) -> dict:
                return {
                    "pid": pid,
                    "ppid": ppid,
                    "pgrp": pgrp,
                    "session": pgrp,
                    "start_time": start_time,
                    "executable": executable,
                    "cwd": str(cwd),
                    "command": executable,
                }

            root_pid = runner_pid + 10_000_002
            child_pid = runner_pid + 10_000_003
            shared_pgrp = root_pid
            launch = {
                ancestor_pid: row(ancestor_pid, 1, shared_pgrp, 1, "cargo", directory),
                runner_pid: row(runner_pid, ancestor_pid, shared_pgrp, 2, "python3", directory),
                root_pid: row(root_pid, runner_pid, shared_pgrp, 3, "python3", directory),
                child_pid: row(child_pid, root_pid, shared_pgrp, 4, "python3", directory),
            }
            ownership = runner.new_ownership(root_pid, directory, target, launch, launch)

            owned = runner.owned_processes(root_pid, launch, ownership)
            owned_pids = {item["pid"] for item in owned}
            self.assertEqual(owned_pids, {root_pid, child_pid})
            builders = runner.owned_builders(root_pid, directory, target, launch, ownership)
            self.assertNotIn(ancestor_pid, {item["pid"] for item in builders})

    def test_run_step_reports_fake_child_failure_without_self_attribution(self) -> None:
        """A genuine failed child remains a failed row, not a runner leak."""
        with tempfile.TemporaryDirectory(prefix="safe-gate-failed-child-") as raw:
            directory = Path(raw)
            root = directory / "source"
            root.mkdir()
            target = directory / "target"
            target.mkdir()
            script = root / "fail.py"
            script.write_text("import sys; print('intentional failure', flush=True); sys.exit(7)\n", encoding="utf-8")
            source = {"commit": "fixture", "status": "", "branch": "fixture", "cargo_locks": {}}
            step = runner.direct("fake-failure", [sys.executable, str(script)])

            with (
                mock.patch.object(runner, "preflight", return_value={}),
                mock.patch.object(runner, "source_state", return_value=source),
            ):
                result = runner.run_step(step, root, target, directory, source, 0, 1, "fixture")

            self.assertEqual(result["exit"], 7)
            self.assertIn("intentional failure", (directory / "00-fake-failure-attempt-fixture.log").read_text(encoding="utf-8"))

    def test_unexpected_guard_exception_kills_fake_command_and_records_failure(self) -> None:
        with tempfile.TemporaryDirectory(prefix="safe-gate-guard-exception-") as raw:
            directory = Path(raw)
            target = directory / "target"
            target.mkdir()
            runner.enable_child_subreaper()
            prelaunch = runner.process_table()
            process = subprocess.Popen(
                [sys.executable, "-c", "import time; time.sleep(5)"],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                start_new_session=True,
            )
            ownership = runner.new_ownership(process.pid, directory, target, prelaunch, runner.process_table())
            stop = threading.Event()
            overflow: list[dict] = []
            errors: list[dict] = []
            stats: dict = {}
            guard = threading.Thread(
                target=lambda: runner._run_target_guard(
                    process, target, stop, overflow, errors, ownership, stats
                )
            )
            try:
                with mock.patch.object(runner, "_watch_target", side_effect=RuntimeError("injected monitor failure")):
                    guard.start()
                    guard.join(timeout=3)
                self.assertFalse(guard.is_alive())
                self.assertEqual(errors[0]["reason"], "target guard exception")
                self.assertIn("injected monitor failure", errors[0]["error"])
                process.wait(timeout=2)
                self.assertIsNotNone(process.returncode)
            finally:
                stop.set()
                if process.poll() is None:
                    runner._terminate_owned_tree(process.pid, ownership, "test-cleanup")
                process.wait()
                guard.join(timeout=2)

    def test_gate_runner_guard_records_persistent_measurement_failure(self) -> None:
        with tempfile.TemporaryDirectory(prefix="safe-gate-du-failure-") as raw:
            directory = Path(raw)
            target = directory / "target"
            target.mkdir()
            process = subprocess.Popen(
                [sys.executable, "-c", "import time; time.sleep(5)"],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                start_new_session=True,
            )
            stop = threading.Event()
            overflow: list[dict] = []
            stats: dict = {}
            guard = threading.Thread(
                target=lambda: stats.update(runner._watch_target(process, target, stop, overflow))
            )
            try:
                with mock.patch.object(runner, "target_bytes", side_effect=[0, RuntimeError("du race")]):
                    guard.start()
                    guard.join(timeout=3)
                self.assertFalse(guard.is_alive())
                self.assertEqual(overflow[0]["reason"], "target measurement failed")
                self.assertIn("du race", overflow[0]["error"])
                process.wait(timeout=2)
            finally:
                stop.set()
                if process.poll() is None:
                    runner._terminate_process_group(process)
                process.wait()
                guard.join(timeout=2)

    def test_gate_runner_live_guard_reaps_setsid_child(self) -> None:
        with tempfile.TemporaryDirectory(prefix="safe-gate-live-escaped-") as raw:
            directory = Path(raw)
            target = directory / "target"
            script = directory / "spawn_escaped.py"
            growth = target / "debug" / "escaped-growth.bin"
            marker = directory / "escaped-marker"
            script.write_text(
                textwrap.dedent(
                    """
                    import os
                    from pathlib import Path
                    import subprocess
                    import sys

                    target = Path(sys.argv[1])
                    marker = Path(sys.argv[2])
                    child = "\\n".join((
                        "import os, pathlib, sys, time",
                        "os.setsid()",
                        "marker = pathlib.Path(sys.argv[2])",
                        "marker.write_text('started')",
                        "path = pathlib.Path(sys.argv[1])",
                        "path.parent.mkdir(parents=True, exist_ok=True)",
                        "deadline = time.monotonic() + 0.55",
                        "while time.monotonic() < deadline:",
                        "    with path.open('ab') as stream:",
                        "        stream.write(b'x' * 4096)",
                        "    time.sleep(0.01)",
                        "marker.write_text('done')",
                    ))
                    subprocess.Popen(
                        [sys.executable, "-c", child, str(target / "debug" / "escaped-growth.bin"), str(marker)],
                        stdout=subprocess.DEVNULL,
                        stderr=subprocess.DEVNULL,
                    )
                    """
                ),
                encoding="utf-8",
            )
            runner.enable_child_subreaper()
            previous_live = runner.LIVE_TARGET
            initial = runner.target_bytes(target)
            runner.LIVE_TARGET = initial + 32_000
            prelaunch = runner.process_table()
            process = subprocess.Popen(
                [sys.executable, str(script), str(target), str(marker)],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                start_new_session=True,
            )
            ownership = runner.new_ownership(process.pid, directory, target, prelaunch, runner.process_table())
            monitor = threading.Thread(target=runner.capture_step_identities, args=(process, ownership))
            stop = threading.Event()
            overflow: list[dict] = []
            stats: dict = {}
            guard = threading.Thread(
                target=lambda: stats.update(runner._watch_target(process, target, stop, overflow, ownership))
            )
            monitor.start()
            guard.start()
            try:
                guard.join(timeout=5)
                self.assertFalse(guard.is_alive())
                process.wait(timeout=2)
                monitor.join(timeout=2)
                self.assertTrue(overflow)
                self.assertFalse(marker.exists() and marker.read_text(encoding="utf-8") == "done")
                time.sleep(0.8)
                self.assertFalse(marker.exists() and marker.read_text(encoding="utf-8") == "done")
                self.assertEqual(runner._owned_identity_rows(process.pid, ownership), [])
            finally:
                stop.set()
                runner._terminate_owned_tree(process.pid, ownership, "test-cleanup")
                if process.poll() is None:
                    process.wait()
                monitor.join(timeout=2)
                guard.join(timeout=2)
                runner.LIVE_TARGET = previous_live

    def test_gate_post_exit_timeout_reaps_setsid_child(self) -> None:
        with tempfile.TemporaryDirectory(prefix="safe-gate-post-exit-escaped-") as raw:
            directory = Path(raw)
            target = directory / "target"
            script = directory / "spawn_hung.py"
            marker = directory / "post-exit-marker"
            script.write_text(
                textwrap.dedent(
                    """
                    import os
                    from pathlib import Path
                    import subprocess
                    import sys

                    marker = Path(sys.argv[1])
                    child = "\\n".join((
                        "import os, pathlib, sys, time",
                        "os.setsid()",
                        "time.sleep(0.55)",
                        "pathlib.Path(sys.argv[1]).write_text('late')",
                    ))
                    subprocess.Popen(
                        [sys.executable, "-c", child, str(marker)],
                        stdout=subprocess.DEVNULL,
                        stderr=subprocess.DEVNULL,
                    )
                    """
                ),
                encoding="utf-8",
            )
            runner.enable_child_subreaper()
            previous_timeout = runner.POST_EXIT_TIMEOUT_SECONDS
            previous_grace = runner.PROCESS_KILL_GRACE_SECONDS
            runner.POST_EXIT_TIMEOUT_SECONDS = 0.1
            runner.PROCESS_KILL_GRACE_SECONDS = 0.1
            prelaunch = runner.process_table()
            process = subprocess.Popen(
                [sys.executable, str(script), str(marker)],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                start_new_session=True,
            )
            ownership = runner.new_ownership(process.pid, directory, target, prelaunch, runner.process_table())
            monitor = threading.Thread(target=runner.capture_step_identities, args=(process, ownership))
            monitor.start()
            try:
                process.wait(timeout=2)
                monitor.join(timeout=2)
                with self.assertRaisesRegex(RuntimeError, "owned child"):
                    runner.post_exit_guard(process.pid, directory, target, ownership)
                time.sleep(0.7)
                self.assertFalse(marker.exists())
                self.assertEqual(runner._owned_identity_rows(process.pid, ownership), [])
            finally:
                runner._terminate_owned_tree(process.pid, ownership, "test-cleanup")
                if process.poll() is None:
                    process.wait()
                monitor.join(timeout=2)
                runner.POST_EXIT_TIMEOUT_SECONDS = previous_timeout
                runner.PROCESS_KILL_GRACE_SECONDS = previous_grace


if __name__ == "__main__":
    unittest.main()
