"""Fake-boundary tests for the final DOCX Ink launcher.

These tests never invoke systemd or a qualification binary.  The fake
``systemd-run`` preserves the wrapper's ``TMPDIR`` contract and the fake
``systemctl`` records an owned-unit stop, which lets the lifecycle and
resource checks run in an ordinary test process.
"""

from __future__ import annotations

import os
from pathlib import Path
import stat
import tempfile
import unittest
from unittest import mock

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

    def test_argv_binds_memory_swap_deadline_and_filesystem_limit(self) -> None:
        runner, _output, tmp, _logs = self._runner(deadline=17.5)
        argv = runner.systemd_argv(["/bin/true"])
        self.assertIn(f"MemoryMax={MEMORY_MAX_BYTES}", argv)
        self.assertIn(f"MemorySwapMax={SWAP_MAX_BYTES}", argv)
        self.assertIn("RuntimeMaxSec=17.5", argv)
        self.assertIn("KillMode=control-group", argv)
        self.assertIn("OOMPolicy=kill", argv)
        self.assertIn(f"LimitFSIZE={runner.limits.output_bytes}", argv)
        self.assertIn(f"--setenv=TMPDIR={tmp}", argv)

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
        self.assertFalse(self.fake_systemctl_log.exists())
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
