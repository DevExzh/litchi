"""Fake-boundary tests for the final DOCX Ink launcher.

These tests never invoke systemd or a qualification binary.  The fake
``systemd-run`` preserves the wrapper's ``TMPDIR`` contract and the fake
``systemctl`` records an owned-unit stop, which lets the lifecycle and
resource checks run in an ordinary test process.
"""

from __future__ import annotations

from pathlib import Path
import stat
import subprocess
import sys
import tempfile
import unittest

from tools.ink_safe_runner import (
    BoundaryError,
    Limits,
    MEMORY_MAX_BYTES,
    OUTPUT_MAX_BYTES,
    SafeInkCommand,
    SWAP_MAX_BYTES,
)


class SafeInkRunnerTests(unittest.TestCase):
    def setUp(self) -> None:
        self.tempdir = tempfile.TemporaryDirectory(prefix="ink-safe-runner-")
        self.root = Path(self.tempdir.name)
        self.fake_systemd = self.root / "fake-systemd-run"
        self.fake_systemctl_log = self.root / "systemctl.log"
        self.fake_systemctl = self.root / "fake-systemctl"
        self.stop_coordination = self.root / "stop-coordination"
        self.stop_marker = self.root / "stop.marker"
        self.done_marker = self.root / "done.marker"
        self.fake_systemd.write_text(
            "#!/bin/sh\n"
            "set -eu\n"
            "tmp=\n"
            "while [ \"$#\" -gt 0 ]; do\n"
            "  case \"$1\" in\n"
            "    --setenv=TMPDIR=*) tmp=${1#--setenv=TMPDIR=} ;;\n"
            "    --) shift; export TMPDIR=\"$tmp\"; exec \"$@\" ;;\n"
            "  esac\n"
            "  shift\n"
            "done\n"
            "exit 64\n",
            encoding="utf-8",
        )
        self.fake_systemctl.write_text(
            "#!/bin/sh\n"
            "set -eu\n"
            f"printf '%s\\n' \"$*\" >> {str(self.fake_systemctl_log)!r}\n"
            f"if [ -f {str(self.stop_coordination)!r} ]; then\n"
            f"  : > {str(self.stop_marker)!r}\n"
            f"  while [ ! -f {str(self.done_marker)!r} ]; do :; done\n"
            "fi\n"
            "case \" $* \" in\n"
            "  *' show '*)\n"
            "    printf '%s\\n' 'ActiveState=inactive' 'SubState=dead' 'Result=success' "
            "'ControlGroup=' 'MemoryCurrent=[not set]' 'MemoryPeak=[not set]' 'MainPID=0' "
            "'ExecMainStatus=0'\n"
            "    ;;\n"
            "esac\n"
            "exit 0\n",
            encoding="utf-8",
        )
        for path in (self.fake_systemd, self.fake_systemctl):
            path.chmod(path.stat().st_mode | stat.S_IXUSR)

    def tearDown(self) -> None:
        self.tempdir.cleanup()

    def _runner(self, *, deadline: float = 2.0, output_cap: int = 1_000_000,
                tmp_cap: int = 1_000_000) -> tuple[SafeInkCommand, Path, Path, Path]:
        output = self.root / "output"
        tmp = self.root / "tmp"
        logs = self.root / "logs"
        runner = SafeInkCommand(
            phase="allocation",
            output=output,
            tmp=tmp,
            log_dir=logs,
            unit="ink-test-unit",
            limits=Limits(deadline, output_bytes=output_cap, tmp_bytes=tmp_cap),
            systemd_run=str(self.fake_systemd),
            systemctl=str(self.fake_systemctl),
        )
        return runner, output, tmp, logs

    def test_argv_binds_systemd_policy_and_child_file_limit(self) -> None:
        runner, _output, tmp, _logs = self._runner(deadline=17.5)
        argv = runner.systemd_argv(["/bin/true"])
        self.assertIn(f"--property=MemoryMax={MEMORY_MAX_BYTES}", argv)
        self.assertIn(f"--property=MemorySwapMax={SWAP_MAX_BYTES}", argv)
        self.assertIn("--property=RuntimeMaxSec=17.5", argv)
        self.assertIn("--property=MemoryAccounting=yes", argv)
        self.assertIn("--property=TasksMax=4096", argv)
        self.assertIn("--property=KillMode=control-group", argv)
        self.assertIn("--ink-child", argv)
        self.assertIn(f"--fsize={runner.limits.output_bytes}", argv)
        self.assertIn(f"--setenv=TMPDIR={tmp}", argv)
        self.assertNotIn("--wait", argv)
        self.assertNotIn("--collect", argv)
        self.assertNotIn(f"LimitFSIZE={runner.limits.output_bytes}", argv)
        self.assertEqual(runner.unit, "ink-test-unit.scope")

    def test_limits_cannot_be_raised_above_policy(self) -> None:
        with self.assertRaises(BoundaryError):
            Limits(1, output_bytes=OUTPUT_MAX_BYTES + 1)

    def test_normal_lifecycle_passes_tmpdir_and_cleans_owned_tmp(self) -> None:
        runner, output, tmp, logs = self._runner()
        command = [
            "/bin/sh",
            "-c",
            f"mkdir -p {str(output)!r}; printf payload > \"$TMPDIR/work\"; "
            f"printf output > {str(output / 'result')!r}",
        ]
        receipt = runner.run(command)
        self.assertEqual(receipt["status"], "complete")
        self.assertIsNone(receipt["violation"])
        self.assertFalse(tmp.exists())
        self.assertEqual((output / "result").read_text(encoding="utf-8"), "output")
        self.assertTrue(self.fake_systemctl_log.exists())
        systemctl_log = self.fake_systemctl_log.read_text(encoding="utf-8")
        self.assertIn("show", systemctl_log)
        self.assertIn("ink-test-unit.scope", systemctl_log)
        self.assertTrue(receipt["post_unit"]["unit_gone"])
        self.assertTrue((logs / "outer-receipt.json").exists())

    def test_tmp_cap_stops_owned_unit_and_cleans_tmp(self) -> None:
        runner, _output, tmp, _logs = self._runner(tmp_cap=4)
        command = [
            "/bin/sh",
            "-c",
            "dd if=/dev/zero of=\"$TMPDIR/grow\" bs=1024 count=16 >/dev/null; sleep 5",
        ]
        receipt = runner.run(command)
        self.assertEqual(receipt["status"], "failed")
        self.assertIn("TMP cap exceeded", receipt["violation"])
        self.assertFalse(tmp.exists())
        self.assertTrue(self.fake_systemctl_log.exists())

    def test_output_cap_stops_owned_unit(self) -> None:
        runner, output, _tmp, _logs = self._runner(output_cap=4096)
        command = [
            "/bin/sh",
            "-c",
            f"mkdir -p {str(output)!r}; "
            f"dd if=/dev/zero of={str(output / 'a')!r} bs=4096 count=1 >/dev/null; "
            f"dd if=/dev/zero of={str(output / 'b')!r} bs=4096 count=1 >/dev/null; sleep 5",
        ]
        receipt = runner.run(command)
        self.assertEqual(receipt["status"], "failed")
        self.assertIn("output cap exceeded", receipt["violation"])
        self.assertTrue(self.fake_systemctl_log.exists())

    def test_fast_exit_output_cap_is_failed(self) -> None:
        runner, output, _tmp, _logs = self._runner(output_cap=4096)
        self.stop_coordination.touch()
        command = [
            "/bin/sh",
            "-c",
            f"mkdir -p {str(output)!r}; "
            f"dd if=/dev/zero of={str(output / 'a')!r} bs=4096 count=1 >/dev/null; "
            f"dd if=/dev/zero of={str(output / 'b')!r} bs=4096 count=1 >/dev/null; "
            f"while [ ! -f {str(self.stop_marker)!r} ]; do :; done; "
            f": > {str(self.done_marker)!r}",
        ]
        receipt = runner.run(command)
        self.assertEqual(receipt["status"], "failed")
        self.assertIn("output cap exceeded", receipt["violation"])
        self.assertEqual(receipt["returncode"], 0)
        self.assertIsNotNone(receipt["cleanup"])

    def test_fast_exit_tmp_cap_is_failed(self) -> None:
        runner, _output, tmp, _logs = self._runner(tmp_cap=4096)
        self.stop_coordination.touch()
        command = [
            "/bin/sh",
            "-c",
            "dd if=/dev/zero of=\"$TMPDIR/a\" bs=4096 count=1 >/dev/null; "
            "dd if=/dev/zero of=\"$TMPDIR/b\" bs=4096 count=1 >/dev/null; "
            f"while [ ! -f {str(self.stop_marker)!r} ]; do :; done; "
            f": > {str(self.done_marker)!r}",
        ]
        receipt = runner.run(command)
        self.assertEqual(receipt["status"], "failed")
        self.assertIn("TMP cap exceeded", receipt["violation"])
        self.assertEqual(receipt["returncode"], 0)
        self.assertIsNotNone(receipt["cleanup"])
        self.assertFalse(tmp.exists())

    def test_exact_output_and_tmp_caps_are_accepted(self) -> None:
        runner, output, tmp, _logs = self._runner(output_cap=4096, tmp_cap=4096)
        command = [
            "/bin/sh",
            "-c",
            f"mkdir -p {str(output)!r}; "
            f"dd if=/dev/zero of={str(output / 'a')!r} bs=4096 count=1 >/dev/null; "
            "dd if=/dev/zero of=\"$TMPDIR/a\" bs=4096 count=1 >/dev/null",
        ]
        receipt = runner.run(command)
        self.assertEqual(receipt["status"], "complete")
        self.assertIsNone(receipt["violation"])
        self.assertEqual(receipt["output"]["bytes"], 4096)
        self.assertEqual(receipt["tmp"]["bytes_before_cleanup"], 4096)
        self.assertFalse(tmp.exists())

    def test_documented_phase_first_cli_runs_with_fake_systemd(self) -> None:
        output = self.root / "cli-output"
        tmp = self.root / "cli-tmp"
        logs = self.root / "cli-logs"
        script = Path(__file__).with_name("ink_safe_runner.py")
        result = subprocess.run(
            [
                sys.executable,
                str(script),
                "allocation",
                "--output",
                str(output),
                "--tmp",
                str(tmp),
                "--log-dir",
                str(logs),
                "--unit",
                "ink-cli-test-unit",
                "--deadline",
                "2",
                "--output-cap",
                "1000000",
                "--tmp-cap",
                "1000000",
                "--systemd-run",
                str(self.fake_systemd),
                "--systemctl",
                str(self.fake_systemctl),
                "--",
                "/bin/true",
            ],
            capture_output=True,
            text=True,
            check=False,
        )
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertTrue((logs / "outer-receipt.json").is_file())
        self.assertFalse(tmp.exists())

    def test_deadline_stops_owned_unit(self) -> None:
        runner, _output, tmp, _logs = self._runner(deadline=0.1)
        receipt = runner.run(["/bin/sh", "-c", "sleep 5"])
        self.assertEqual(receipt["status"], "failed")
        self.assertIn("deadline exceeded", receipt["violation"])
        self.assertFalse(tmp.exists())
        self.assertTrue(self.fake_systemctl_log.exists())

    def test_output_and_tmp_paths_must_be_disjoint(self) -> None:
        output = self.root / "output"
        with self.assertRaises(BoundaryError):
            SafeInkCommand(
                phase="allocation",
                output=output,
                tmp=output / "tmp",
                log_dir=self.root / "logs",
                unit="ink-test-unit",
                limits=Limits(1),
                systemd_run=str(self.fake_systemd),
                systemctl=str(self.fake_systemctl),
            )


if __name__ == "__main__":
    unittest.main()
