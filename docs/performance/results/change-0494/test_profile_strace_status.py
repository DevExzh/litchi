#!/usr/bin/env python3
"""Focused tests for the future strace observer failure classification."""

from __future__ import annotations

from pathlib import Path
import sys
import tempfile
import unittest

sys.path.insert(0, str(Path(__file__).resolve().parent))
import profile as profiler  # noqa: E402


class StraceStatusTests(unittest.TestCase):
    def _status(self, diagnostic: str, **process: object) -> tuple[str, str | None]:
        with tempfile.TemporaryDirectory() as directory:
            stderr = Path(directory) / "strace.stderr"
            stderr.write_text(diagnostic, encoding="utf-8")
            return profiler._strace_failure_status(process, stderr)

    def test_ptrace_setup_denial_is_unavailable(self) -> None:
        status, reason = self._status(
            "strace: attach: ptrace(PTRACE_SEIZE, 123): Operation not permitted\n",
            exit_code=1,
            timed_out=False,
        )
        self.assertEqual(status, "unavailable")
        self.assertIn("ptrace", reason or "")

    def test_target_permission_error_remains_failed(self) -> None:
        status, reason = self._status(
            "target: permission denied\n",
            exit_code=1,
            timed_out=False,
        )
        self.assertEqual(status, "failed")
        self.assertIn("nonzero", reason or "")

    def test_missing_strace_process_is_unavailable(self) -> None:
        status, reason = self._status(
            "",
            launch_error="FileNotFoundError: strace",
            exit_code=None,
            timed_out=False,
        )
        self.assertEqual(status, "unavailable")
        self.assertIn("launched", reason or "")


if __name__ == "__main__":
    raise SystemExit(unittest.main())
