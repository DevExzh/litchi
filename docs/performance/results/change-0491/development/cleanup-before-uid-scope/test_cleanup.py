#!/usr/bin/env python3
"""Isolated custody tests for ``cleanup.py``.

The tests use temporary evidence/cache/process trees.  They never point the
destructive path at the real 0491 roots and never modify the repository's
evidence bundle.
"""

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


sys.dont_write_bytecode = True

_DRIVER = Path(__file__).with_name("cleanup.py").resolve()
_SPEC = importlib.util.spec_from_file_location("cleanup_0491_under_test", _DRIVER)
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
        self.holder = tempfile.TemporaryDirectory(prefix="litchi-cleanup-test-")
        base = Path(self.holder.name)
        self.root = base / "evidence"
        self.temp = base / "temporary"
        self.target = base / "cargo-target"
        self.proc = base / "proc"
        self.root.mkdir()
        self.temp.mkdir()
        self.target.mkdir()
        self.proc.mkdir()
        (self.root / "validation").mkdir()
        self._make_receipts()

    def tearDown(self) -> None:
        self.holder.cleanup()

    def _make_receipts(self) -> None:
        final = self.temp / "final5"
        for role, name in cleanup.BUILD_ROLES.items():
            role_dir = final / role
            role_dir.mkdir(parents=True)
            binary = role_dir / name
            binary.write_bytes((role + " benchmark\n").encode())
            binary.chmod(binary.stat().st_mode | stat.S_IXUSR)
            original = self.target / "release" / name
            original.parent.mkdir(parents=True, exist_ok=True)
            original.write_bytes(binary.read_bytes())
            original.chmod(original.stat().st_mode | stat.S_IXUSR)
            metadata = {"bytes": binary.stat().st_size, "sha256": _sha(binary)}
            original_metadata = {"bytes": original.stat().st_size, "sha256": _sha(original)}
            _write_json(
                self.root / f"build-{role}.json",
                {
                    "schema": cleanup.BUILD_SCHEMA,
                    "version": cleanup.VERSION,
                    "attempt": cleanup.FINAL_ATTEMPT,
                    "role": role,
                    "source_unchanged": True,
                    "source_before": {"sha256": "a" * 64},
                    "source_after": {"sha256": "a" * 64},
                    "binary": {
                        "path": str(binary),
                        **metadata,
                        "executable": True,
                    },
                    "original_binary": {
                        "path": str(original),
                        **original_metadata,
                        "executable": True,
                    },
                },
            )

        started = self.root / "validation" / "isolated.started.json"
        result = self.root / "validation" / "isolated.json"
        _write_json(
            started,
            {"schema": "docx-stream-append-gate-v1", "started_utc": "now"},
        )
        _write_json(
            result,
            {
                "schema": "docx-stream-append-gate-v1",
                "exit_code": 0,
                "finished_utc": "now",
            },
        )

    def _plan(self) -> dict[str, object]:
        return cleanup.plan_cleanup(
            root=self.root,
            temp=self.temp,
            target=self.target,
            proc_root=self.proc,
            self_pid=999_999,
        )

    def test_plan_is_read_only_and_reports_allocated_bytes(self) -> None:
        disposable = self.temp / "draft"
        disposable.mkdir()
        (disposable / "payload").write_bytes(b"temporary payload")
        (self.target / "incremental").mkdir()
        (self.target / "incremental" / "metadata").write_bytes(b"metadata")

        plan = self._plan()

        self.assertTrue(disposable.exists())
        self.assertTrue(self.target.exists())
        self.assertEqual({child.name for child in self.temp.iterdir()}, {"draft", "final5"})
        removed = plan["removed_stats"]
        self.assertTrue(all(isinstance(item["allocated_bytes"], int) for item in removed))
        self.assertGreaterEqual(plan["process_before"]["scanned_processes"], 0)

    def test_isolated_execution_removes_only_disposable_roots_and_verify_is_empty(self) -> None:
        (self.temp / "old").mkdir()
        (self.temp / "old" / "scratch").write_bytes(b"scratch")
        (self.target / "fingerprint").write_bytes(b"cargo output")

        receipt = cleanup.execute_cleanup(self._plan())

        self.assertEqual(receipt["status"], "pass")
        self.assertFalse((self.temp / "old").exists())
        self.assertFalse(self.target.exists())
        self.assertEqual({child.name for child in self.temp.iterdir()}, {"final5"})
        for role, name in cleanup.BUILD_ROLES.items():
            retained = self.temp / "final5" / role / name
            self.assertTrue(retained.is_file())
            self.assertEqual(receipt["retained_binaries"][list(cleanup.BUILD_ROLES).index(role)]["sha256"], _sha(retained))

        proof = cleanup.verify(
            root=self.root,
            temp=self.temp,
            target=self.target,
            proc_root=self.proc,
        )
        self.assertEqual(proof["status"], "pass")
        self.assertEqual(proof["remaining"], [])
        self.assertTrue((self.root / "cleanup.json").is_file())

    def test_existing_cleanup_receipt_is_never_replaced(self) -> None:
        (self.temp / "old").mkdir()
        (self.root / "cleanup.json").write_text("old\n", encoding="utf-8")

        with self.assertRaises(cleanup.CleanupError):
            self._plan()
        self.assertEqual((self.root / "cleanup.json").read_text(encoding="utf-8"), "old\n")
        self.assertTrue((self.temp / "old").exists())

    def test_binary_hash_drift_aborts_before_any_removal(self) -> None:
        (self.temp / "old").mkdir()
        binary = self.temp / "final5" / "normal" / cleanup.BUILD_ROLES["normal"]
        binary.write_bytes(b"changed")

        with self.assertRaises(cleanup.CleanupError):
            self._plan()
        self.assertTrue((self.temp / "old").exists())
        self.assertTrue(self.target.exists())

    def test_symlink_candidate_is_refused(self) -> None:
        outside = Path(self.holder.name) / "outside"
        outside.mkdir()
        (outside / "secret").write_bytes(b"do not remove")
        (self.temp / "link").symlink_to(outside, target_is_directory=True)

        with self.assertRaises(cleanup.CleanupError):
            self._plan()
        self.assertTrue((outside / "secret").exists())
        self.assertTrue((self.temp / "link").is_symlink())

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

    def test_cwd_exe_and_fd_references_abort(self) -> None:
        candidate = self.temp / "old"
        candidate.mkdir()
        (candidate / "data").write_bytes(b"data")
        self._fake_process(123, cwd=candidate)

        with self.assertRaisesRegex(cleanup.CleanupError, "active process references"):
            self._plan()
        self.assertTrue(candidate.exists())

    def test_executable_and_file_descriptor_references_are_reported(self) -> None:
        candidate = self.temp / "old"
        candidate.mkdir()
        executable = candidate / "worker"
        executable.write_bytes(b"worker")
        executable.chmod(executable.stat().st_mode | stat.S_IXUSR)
        opened = candidate / "opened"
        opened.write_bytes(b"opened")
        self._fake_process(125, exe=executable, fd=opened)

        audit = cleanup._process_audit(
            [candidate],
            [self.temp, self.target],
            self.root,
            proc_root=self.proc,
            self_pid=999_999,
        )

        self.assertFalse(audit["safe"])
        self.assertEqual(
            {item["reference"] for item in audit["candidate_references"]},
            {"exe", "fd:3"},
        )

    def test_live_root_capture_script_aborts_even_without_candidate_reference(self) -> None:
        candidate = self.temp / "old"
        candidate.mkdir()
        script = self.root / "provider_matrix.py"
        script.write_text("# fixture\n", encoding="utf-8")
        python = Path(sys.executable)
        self._fake_process(
            124,
            cwd=self.root,
            exe=python,
            command=[str(python), "provider_matrix.py"],
        )

        with self.assertRaisesRegex(cleanup.CleanupError, "live root gate/capture"):
            self._plan()
        self.assertTrue(candidate.exists())

    def test_destructive_cli_rejects_noncanonical_roots(self) -> None:
        (self.temp / "old").mkdir()

        status = cleanup.main(
            [
                "--root",
                str(self.root),
                "--temp-root",
                str(self.temp),
                "--target-root",
                str(self.target),
                "--proc-root",
                str(self.proc),
            ]
        )

        self.assertEqual(status, 1)
        self.assertTrue((self.temp / "old").exists())
        self.assertFalse((self.root / "cleanup.json").exists())


if __name__ == "__main__":
    unittest.main()
