#!/usr/bin/env python3
"""Executable lifecycle tests for the fail-closed OPC timing guard.

These tests intentionally use a tiny Python child instead of a Cargo/native
binary.  They exercise ownership, subreaper adoption, zombie reaping, target
identity, restricted metadata classification, and the complete post-run host
snapshot without consuming a build or timing slot.
"""

from __future__ import annotations

import importlib.util
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import time
import unittest


ROOT = Path(__file__).resolve().parents[1]
GUARD_PATH = ROOT / "docs/unified-ops/results/0362umf/opc_read_ahead_timing.py"
SPEC = importlib.util.spec_from_file_location("opc_read_ahead_timing", GUARD_PATH)
if SPEC is None or SPEC.loader is None:
    raise RuntimeError(f"cannot import guard: {GUARD_PATH}")
GUARD = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(GUARD)


FAKE_CHILD = r'''
import os
from pathlib import Path
import signal
import sys
import time

mode, marker = sys.argv[1], Path(sys.argv[2])
if mode == "sleep":
    time.sleep(60)
elif mode == "escape":
    child = os.fork()
    if child == 0:
        os.setsid()
        marker.write_text(str(os.getpid()), encoding="ascii")
        signal.signal(signal.SIGTERM, signal.SIG_IGN)
        time.sleep(60)
        os._exit(0)
    os._exit(0)
else:
    raise SystemExit(f"unknown fake mode: {mode}")
'''


def spawn_fake(root: Path, mode: str) -> tuple[subprocess.Popen[bytes], int, dict[int, int], set[int]]:
    script = root / "fake_child.py"
    if not script.exists():
        script.write_text(FAKE_CHILD, encoding="utf-8")
        script.chmod(0o700)
    marker = root / f"{mode}.pid"
    before = GUARD.proc_snapshot()
    process = subprocess.Popen(
        [sys.executable, str(script), mode, str(marker)],
        stdin=subprocess.DEVNULL,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        start_new_session=True,
    )
    pgid = os.getpgid(process.pid)
    entry = GUARD.read_proc(process.pid)
    if entry is None:
        raise AssertionError("fake child disappeared before identity capture")
    owned = {process.pid: int(entry["starttime"])}
    return process, pgid, owned, set(before)


def wait_marker_pid(marker: Path, deadline: float) -> int:
    """Wait for the escaped child to publish a complete identity marker."""
    last_value = ""
    while time.monotonic() < deadline:
        try:
            last_value = marker.read_text(encoding="ascii").strip()
        except FileNotFoundError:
            last_value = ""
        if last_value:
            try:
                value = int(last_value)
            except ValueError:
                value = 0
            if value > 0:
                return value
        time.sleep(0.01)
    raise AssertionError(f"escaped fake child did not publish a PID: {last_value!r}")


class OpcReadAheadGuardTests(unittest.TestCase):
    def test_monitor_failure_still_waits_and_reaps_direct_child(self) -> None:
        with tempfile.TemporaryDirectory(prefix="opc-guard-lifecycle-", dir="/tmp") as raw:
            root = Path(raw)
            process, pgid, owned, _ = spawn_fake(root, "sleep")
            try:
                with self.assertRaises(RuntimeError):
                    try:
                        raise RuntimeError("simulated monitor failure")
                    except RuntimeError:
                        GUARD.ensure_child_reaped(process, owned, pgid, False)
                        raise
                self.assertIsNotNone(process.returncode)
                self.assertEqual(process.wait(timeout=1), process.returncode)
                with self.assertRaises(ProcessLookupError):
                    os.kill(process.pid, 0)
            finally:
                if process.poll() is None:
                    GUARD.ensure_child_reaped(process, owned, pgid, True)

    def test_adopted_zombie_is_reaped_after_process_group_escape(self) -> None:
        previous_subreaper = GUARD.set_child_subreaper(True)
        try:
            with tempfile.TemporaryDirectory(prefix="opc-guard-escape-", dir="/tmp") as raw:
                root = Path(raw)
                process, pgid, owned, before = spawn_fake(root, "escape")
                marker = root / "escape.pid"
                deadline = time.monotonic() + 2
                escaped_pid = wait_marker_pid(marker, deadline)
                current = GUARD.proc_snapshot()
                owned = GUARD.refresh_owned(current, owned, process.pid, pgid, process.pid)
                for pid, entry in current.items():
                    if pid not in before and int(entry.get("ppid", -1)) == os.getpid():
                        owned[pid] = int(entry["starttime"])
                self.assertIn(escaped_pid, owned)
                GUARD.ensure_child_reaped(process, owned, pgid, False)
                self.assertEqual(process.wait(timeout=1), process.returncode)
                for _ in range(20):
                    if escaped_pid not in GUARD.proc_snapshot():
                        break
                    time.sleep(0.01)
                self.assertNotIn(escaped_pid, GUARD.proc_snapshot())
        finally:
            GUARD.set_child_subreaper(bool(previous_subreaper))

    def test_monitor_snapshot_failure_discovers_escaped_child_for_cleanup(self) -> None:
        previous_subreaper = GUARD.set_child_subreaper(True)
        original_snapshot = GUARD.proc_snapshot
        process: subprocess.Popen[bytes] | None = None
        cleanup_pgid: int | None = None
        cleanup_owned: dict[int, int] = {}
        cleanup_before: set[int] = set()
        escaped_pid: int | None = None
        try:
            with tempfile.TemporaryDirectory(prefix="opc-guard-monitor-failure-", dir="/tmp") as raw:
                root = Path(raw)
                process, pgid, owned, before = spawn_fake(root, "escape")
                cleanup_pgid, cleanup_owned, cleanup_before = pgid, dict(owned), set(before)
                marker = root / "escape.pid"
                deadline = time.monotonic() + 2
                escaped_pid = wait_marker_pid(marker, deadline)

                def failed_snapshot() -> dict[int, dict[str, object]]:
                    raise GUARD.TimingError("synthetic first during-poll failure")

                GUARD.proc_snapshot = failed_snapshot  # type: ignore[assignment]
                GUARD.ensure_child_reaped(
                    process,
                    owned,
                    pgid,
                    False,
                    runner_pid=os.getpid(),
                    preexisting_pids=before,
                )
                GUARD.proc_snapshot = original_snapshot
                self.assertEqual(process.wait(timeout=1), process.returncode)
                self.assertNotIn(escaped_pid, GUARD.proc_snapshot())
        finally:
            GUARD.proc_snapshot = original_snapshot
            if process is not None and cleanup_pgid is not None:
                try:
                    GUARD.ensure_child_reaped(
                        process,
                        cleanup_owned,
                        cleanup_pgid,
                        True,
                        runner_pid=os.getpid(),
                        preexisting_pids=cleanup_before,
                    )
                except (GUARD.TimingError, ProcessLookupError, OSError):
                    pass
            GUARD.set_child_subreaper(bool(previous_subreaper))

    def test_restricted_metadata_requires_nonheavy_nonlitchi_proof(self) -> None:
        self.assertTrue(GUARD.restricted_metadata_is_unrelated("systemd", ["/usr/lib/systemd/systemd"], False))
        self.assertFalse(GUARD.restricted_metadata_is_unrelated("python", [], False))
        self.assertFalse(GUARD.restricted_metadata_is_unrelated("litchi-worker", ["litchi-worker"], False))
        self.assertTrue(GUARD.restricted_metadata_is_unrelated("kthreadd", [], True))

    def test_post_snapshot_and_target_identity_cover_mutation(self) -> None:
        current = {
            9001: {
                "pid": 9001,
                "comm": "rustc",
                "exe": "/usr/bin/rustc",
                "cmdline": ["rustc"],
            }
        }
        violations = GUARD.quiet_violations(
            {},
            current,
            owned={},
            load_before=0.0,
            max_load_delta=100.0,
            previous_cpu=(10, 5),
            current_cpu=(10, 5),
            check_load_delta=False,
            absolute_load_limit=None,
        )
        self.assertTrue(any("unrelated heavy process" in value for value in violations))
        with tempfile.TemporaryDirectory(prefix="opc-guard-target-", dir="/tmp") as raw:
            root = Path(raw)
            target = root / "target"
            target.mkdir()
            payload = target / "payload.bin"
            payload.write_bytes(b"AAAA")
            before = GUARD.target_identity(target, 1024)
            payload.write_bytes(b"BBBB")
            same_size = GUARD.target_identity(target, 1024)
            self.assertEqual(before["bytes"], same_size["bytes"])
            self.assertNotEqual(before["sha256"], same_size["sha256"])
            payload.unlink()
            deleted = GUARD.target_identity(target, 1024)
            self.assertNotEqual(before["sha256"], deleted["sha256"])
            payload.symlink_to(root / "missing")
            with self.assertRaises(GUARD.TimingError):
                GUARD.target_identity(target, 1024)


if __name__ == "__main__":
    unittest.main()
