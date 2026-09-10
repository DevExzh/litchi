#!/usr/bin/env python3
"""Fixture-only custody and safety tests for the 0494 cleanup driver."""

from __future__ import annotations

import hashlib
import importlib.util
import json
import os
from pathlib import Path
import stat
import sys
import tempfile
import unittest
from unittest import mock


sys.dont_write_bytecode = True

_DRIVER = Path(__file__).with_name("cleanup.py").resolve()
_SPEC = importlib.util.spec_from_file_location("cleanup_0494_under_test", _DRIVER)
assert _SPEC is not None and _SPEC.loader is not None
cleanup = importlib.util.module_from_spec(_SPEC)
sys.modules[_SPEC.name] = cleanup
_SPEC.loader.exec_module(cleanup)


def _sha(path: Path) -> str:
    with path.open("rb") as stream:
        return hashlib.file_digest(stream, "sha256").hexdigest()


def _write_json(path: Path, value: object) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(value, sort_keys=True) + "\n", encoding="utf-8")


class CleanupFixture(unittest.TestCase):
    def setUp(self) -> None:
        self.holder = tempfile.TemporaryDirectory(prefix="litchi-cleanup-0494-")
        base = Path(self.holder.name)
        self.root = base / "evidence"
        self.temp = base / "scratch"
        self.target = base / "cargo-target"
        self.proc = base / "proc"
        for path in (self.root, self.temp, self.target, self.proc):
            path.mkdir()
        (self.root / "validation").mkdir()
        self.attempt = "retained-proof"
        self._make_receipts()

    def tearDown(self) -> None:
        self.holder.cleanup()

    def _make_receipts(self) -> None:
        retained = self.temp / self.attempt
        for role, name in cleanup.BUILD_ROLES.items():
            binary = retained / role / name
            binary.parent.mkdir(parents=True, exist_ok=True)
            binary.write_bytes((role + " benchmark\n").encode())
            binary.chmod(binary.stat().st_mode | stat.S_IXUSR)
            original = self.target / "release" / name
            original.parent.mkdir(parents=True, exist_ok=True)
            original.write_bytes(binary.read_bytes())
            original.chmod(original.stat().st_mode | stat.S_IXUSR)
            binary_meta = {"bytes": binary.stat().st_size, "sha256": _sha(binary)}
            original_meta = {"bytes": original.stat().st_size, "sha256": _sha(original)}
            _write_json(self.root / f"build-{role}.json", {
                "schema": cleanup.BUILD_SCHEMA,
                "version": cleanup.VERSION,
                "attempt": self.attempt,
                "role": role,
                "source_unchanged": True,
                "source_before": {"sha256": "a" * 64},
                "source_after": {"sha256": "a" * 64},
                "binary": {"path": str(binary), "executable": True, **binary_meta},
                "original_binary": {"path": str(original), "executable": True, **original_meta},
            })
        started = self.root / "validation" / "fixture.started.json"
        result = self.root / "validation" / "fixture.json"
        _write_json(started, {"schema": "fixture", "started_utc": "now"})
        _write_json(result, {"schema": "fixture", "exit_code": 0, "finished_utc": "now"})

    def _plan(self) -> dict[str, object]:
        return cleanup.plan_cleanup(root=self.root, temp=self.temp, target=self.target, proc_root=self.proc, self_pid=999999)

    def _fake_process(self, pid: int, *, cwd: Path | None = None, exe: Path | None = None, fd: Path | None = None, command: list[str] | None = None) -> None:
        process = self.proc / str(pid)
        (process / "fd").mkdir(parents=True)
        if cwd is not None:
            (process / "cwd").symlink_to(cwd, target_is_directory=True)
        if exe is not None:
            (process / "exe").symlink_to(exe)
        if fd is not None:
            (process / "fd" / "3").symlink_to(fd)
        (process / "cmdline").write_bytes(b"\0".join(item.encode() for item in (command or ["worker"])) + b"\0")

    def test_dry_plan_is_read_only_and_derives_attempt(self) -> None:
        (self.temp / "old").mkdir()
        (self.temp / "old" / "payload").write_bytes(b"scratch")
        (self.target / "incremental").mkdir()
        (self.target / "incremental" / "metadata").write_bytes(b"metadata")
        plan = self._plan()
        self.assertEqual(plan["attempt"], self.attempt)
        self.assertTrue((self.temp / "old").exists())
        self.assertTrue(self.target.exists())
        self.assertTrue(all(isinstance(item["device"], int) for item in plan["removed_stats"]))

    def test_execute_and_verify_leave_exactly_two_authenticated_binaries(self) -> None:
        (self.temp / "old").mkdir()
        (self.temp / "old" / "scratch").write_bytes(b"scratch")
        (self.target / "fingerprint").write_bytes(b"cargo output")
        receipt = cleanup.execute_cleanup(self._plan())
        self.assertEqual(receipt["status"], "pass")
        self.assertFalse((self.temp / "old").exists())
        self.assertFalse(self.target.exists())
        self.assertEqual({child.name for child in self.temp.iterdir()}, {self.attempt})
        for role, name in cleanup.BUILD_ROLES.items():
            self.assertTrue((self.temp / self.attempt / role / name).is_file())
        proof = cleanup.verify(root=self.root, temp=self.temp, target=self.target, proc_root=self.proc)
        self.assertEqual(proof["status"], "pass")
        self.assertEqual(proof["remaining"], [])

    def test_existing_receipt_is_never_replaced(self) -> None:
        (self.temp / "old").mkdir()
        (self.root / "cleanup.json").write_text("preserve\n", encoding="utf-8")
        with self.assertRaises(cleanup.CleanupError):
            self._plan()
        self.assertEqual((self.root / "cleanup.json").read_text(encoding="utf-8"), "preserve\n")
        self.assertTrue((self.temp / "old").exists())

    def test_binary_drift_aborts_before_removal(self) -> None:
        (self.temp / "old").mkdir()
        binary = self.temp / self.attempt / "normal" / cleanup.BUILD_ROLES["normal"]
        binary.write_bytes(b"changed")
        with self.assertRaises(cleanup.CleanupError):
            self._plan()
        self.assertTrue((self.temp / "old").exists())
        self.assertTrue(self.target.exists())

    def test_attempts_must_match_and_retained_layout_has_no_extra_paths(self) -> None:
        value = json.loads((self.root / "build-allocator.json").read_text())
        value["attempt"] = "other"
        _write_json(self.root / "build-allocator.json", value)
        with self.assertRaises(cleanup.CleanupError):
            self._plan()
        value["attempt"] = self.attempt
        _write_json(self.root / "build-allocator.json", value)
        (self.temp / self.attempt / "normal" / "unexpected").write_bytes(b"x")
        with self.assertRaises(cleanup.CleanupError):
            self._plan()

    def test_symlink_and_special_candidates_are_refused(self) -> None:
        outside = Path(self.holder.name) / "outside"
        outside.mkdir()
        (outside / "secret").write_bytes(b"do not remove")
        (self.temp / "link").symlink_to(outside, target_is_directory=True)
        with self.assertRaises(cleanup.CleanupError):
            self._plan()
        self.assertTrue((outside / "secret").exists())

        (self.temp / "link").unlink()
        fifo = self.temp / "fifo"
        os.mkfifo(fifo)
        with self.assertRaises(cleanup.CleanupError):
            self._plan()
        self.assertTrue(fifo.exists())

    def test_cwd_exe_and_fd_references_abort(self) -> None:
        candidate = self.temp / "old"
        candidate.mkdir()
        opened = candidate / "opened"
        opened.write_bytes(b"opened")
        executable = candidate / "worker"
        executable.write_bytes(b"worker")
        executable.chmod(executable.stat().st_mode | stat.S_IXUSR)
        self._fake_process(101, cwd=candidate, exe=executable, fd=opened)
        audit = cleanup._process_audit([candidate], self.root, self.temp, self.target, proc_root=self.proc, self_pid=999999)
        self.assertFalse(audit["safe"])
        self.assertEqual({item["reference"] for item in audit["candidate_references"]}, {"cwd", "exe", "fd:3"})

    def test_live_gate_or_capture_script_aborts_without_candidate_reference(self) -> None:
        candidate = self.temp / "old"
        candidate.mkdir()
        python = Path(sys.executable)
        self._fake_process(102, cwd=self.root, exe=python, command=[str(python), "measure.py"])
        with self.assertRaisesRegex(cleanup.CleanupError, "live 0494 gate/capture"):
            self._plan()
        self.assertTrue(candidate.exists())

    def test_production_cleanup_checks_full_capture_custody(self) -> None:
        with mock.patch("measure.load_builds", side_effect=RuntimeError("forged gate")) as validate:
            with self.assertRaisesRegex(cleanup.CleanupError, "forged gate"):
                cleanup._validate_capture_custody(cleanup.ROOT, cleanup.TEMP, cleanup.TARGET)
            validate.assert_called_once_with()
        with self.assertRaisesRegex(cleanup.CleanupError, "all fixed roots"):
            cleanup._validate_capture_custody(cleanup.ROOT, self.temp, self.target)

    def test_destructive_cli_rejects_fixture_roots(self) -> None:
        (self.temp / "old").mkdir()
        status = cleanup.main(["--root", str(self.root), "--temp-root", str(self.temp), "--target-root", str(self.target), "--proc-root", str(self.proc)])
        self.assertEqual(status, 1)
        self.assertTrue((self.temp / "old").exists())
        self.assertFalse((self.root / "cleanup.json").exists())


if __name__ == "__main__":
    unittest.main()
