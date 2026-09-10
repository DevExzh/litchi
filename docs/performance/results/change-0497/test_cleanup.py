#!/usr/bin/env python3
"""Isolated safety tests for the 0497 cleanup driver.

Every destructive test uses a temporary evidence/scratch/proc fixture.  No
test invokes the fixed production roots.
"""

from __future__ import annotations

import hashlib
import importlib.util
import json
import os
from pathlib import Path
import stat
import shutil
import sys
import tempfile
import unittest
from unittest.mock import patch


sys.dont_write_bytecode = True
_DRIVER = Path(__file__).with_name("cleanup.py").resolve()
_SPEC = importlib.util.spec_from_file_location("cleanup_0497_under_test", _DRIVER)
assert _SPEC is not None and _SPEC.loader is not None
cleanup = importlib.util.module_from_spec(_SPEC)
sys.modules[_SPEC.name] = cleanup
_SPEC.loader.exec_module(cleanup)


def _sha(path: Path) -> str:
    with path.open("rb") as stream:
        return hashlib.file_digest(stream, "sha256").hexdigest()


def _meta(path: Path) -> dict[str, object]:
    return {"path": str(path), "bytes": path.stat().st_size, "sha256": _sha(path)}


def _write_json(path: Path, value: object) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(value, sort_keys=True) + "\n", encoding="utf-8")


class CleanupFixture(unittest.TestCase):
    def setUp(self) -> None:
        self.holder = tempfile.TemporaryDirectory(prefix="litchi-cleanup-0497-")
        base = Path(self.holder.name)
        self.root = base / "evidence"
        self.temp = base / "scratch"
        self.target = self.temp / "target"
        self.fuzz_temp = self.temp / "fuzz-target"
        self.proc = base / "proc"
        for path in (self.root, self.temp, self.target, self.fuzz_temp, self.proc):
            path.mkdir()
        (self.root / "builds").mkdir()
        (self.root / "validation").mkdir()
        (self.root / "fuzz").mkdir()
        (self.root / "build.py").write_text("# fixture build driver\n", encoding="utf-8")
        (self.root / "fuzz.py").write_text("# fixture fuzz driver\n", encoding="utf-8")
        self.receipts = self._make_receipts()

    def tearDown(self) -> None:
        self.holder.cleanup()

    def _make_receipts(self) -> dict[str, Path]:
        result: dict[str, Path] = {}
        for phase in cleanup.PHASES:
            phase_root = self.temp / phase
            phase_root.mkdir()
            (phase_root / "source.txt").write_bytes((phase + " source\n").encode())
            source_manifest_path = self.root / "builds" / f"{phase}-source.json"
            source_manifest = {"source.txt": _meta(phase_root / "source.txt")}
            _write_json(source_manifest_path, source_manifest)
            source_meta = _meta(source_manifest_path)
            phase_receipt: dict[str, object] = {}
            for role in cleanup.BUILD_ROLES:
                retained = self.temp / "retained" / phase / role / cleanup.BINARY_NAME
                retained.parent.mkdir(parents=True, exist_ok=True)
                retained.write_bytes(f"{phase}-{role}-binary\n".encode())
                retained.chmod(retained.stat().st_mode | stat.S_IXUSR)
                gate_path = self.root / "builds" / f"{phase}-candidate1-{role}.json"
                stdout = gate_path.with_suffix(".stdout")
                stderr = gate_path.with_suffix(".stderr")
                stdout.write_bytes(b"")
                stderr.write_bytes(b"")
                driver = _meta(self.root / "build.py")
                gate = {
                    "argv": ["cargo", "build", cleanup.BINARY_NAME],
                    "cwd": str(phase_root),
                    "driver": driver,
                    "environment": {"CARGO_TARGET_DIR": str(self.target)},
                    "exit_code": 0,
                    "finished_ns": 2,
                    "pid": 123,
                    "source_manifest": source_meta,
                    "source_unchanged": True,
                    "started_ns": 1,
                    "stdout": _meta(stdout),
                    "stderr": _meta(stderr),
                }
                _write_json(gate_path.with_suffix(".started.json"), {
                    "argv": gate["argv"],
                    "cwd": gate["cwd"],
                    "driver": driver,
                    "environment": gate["environment"],
                    "source_manifest": source_meta,
                    "started_ns": 1,
                })
                _write_json(gate_path, gate)
                phase_receipt[f"{phase}/{role}"] = {
                    "binary": _meta(retained),
                    "gate": _meta(gate_path),
                    "git_revision": "a" * 40,
                    "source_manifest": source_meta,
                }
            receipt_path = self.root / f"build-{phase}-candidate1.json"
            _write_json(receipt_path, phase_receipt)
            result[phase] = receipt_path

        self._make_fuzz_receipts()

        started = self.root / "validation" / "fixture.started.json"
        terminal = self.root / "validation" / "fixture.json"
        _write_json(started, {"started_ns": 1, "cwd": str(self.temp / "after")})
        _write_json(terminal, {
            "exit_code": 101,
            "finished_ns": 2,
            "source_unchanged": True,
        })
        return result

    def _make_fuzz_receipts(self) -> None:
        evidence = self.root / "fuzz"
        package = self.fuzz_temp / "fuzz"
        target = self.fuzz_temp / "target" / cleanup.FUZZ_TARGET_TRIPLE / "release"
        (self.fuzz_temp / "runs").mkdir()
        (self.fuzz_temp / "tmp").mkdir()
        package.mkdir()
        (package / "source_backed_tail_append_stream.rs").write_bytes(b"target\n")
        (package / "Cargo.toml").write_bytes(b"[package]\nname='fixture'\n")
        (package / "Cargo.lock").write_bytes(b"# fixture lock\n")
        corpus = package / "corpus-start"
        corpus.mkdir()
        (corpus / "seed.bin").write_bytes(b"seed\n")
        target.mkdir(parents=True)
        origin = target / cleanup.FUZZ_BINARY_NAME
        origin.write_bytes(b"asan fuzz binary\n")
        origin.chmod(origin.stat().st_mode | stat.S_IXUSR)
        retained = self.temp / "retained" / "fuzz" / cleanup.FUZZ_BINARY_NAME
        retained.parent.mkdir(parents=True, exist_ok=True)
        retained.write_bytes(origin.read_bytes())
        retained.chmod(retained.stat().st_mode | stat.S_IXUSR)

        old = evidence / "retained-0485"
        (old / "seeds").mkdir(parents=True)
        (old / "fuzz-stream.py").write_bytes(b"old driver\n")
        (old / "generator.json").write_bytes(b"old generator\n")
        (old / "seed-manifest.json").write_bytes(b"old manifest\n")
        (old / "seeds" / "seed.bin").write_bytes(b"old seed\n")
        source_binding = {"fixture_source": "authenticated", "revision": "a" * 40}
        corpus_meta = {"seed.bin": _meta(corpus / "seed.bin")}
        retained_old = {
            "driver": _meta(old / "fuzz-stream.py"),
            "generator": _meta(old / "generator.json"),
            "manifest": _meta(old / "seed-manifest.json"),
            "generator_sha256": "b" * 64,
            "seed_count": 1,
            "seed_inventory": {"seed.bin": _meta(old / "seeds" / "seed.bin")},
        }
        prepared = {
            "schema": "docx-stream-fuzz-prepared-0497-v1",
            "driver": _meta(self.root / "fuzz.py"),
            "target": cleanup.FUZZ_BINARY_NAME,
            "run_seeds": list(cleanup.FUZZ_RUN_SEEDS),
            "paths": {
                "temp_root": str(self.fuzz_temp),
                "package": str(package),
                "cargo_target": str(self.fuzz_temp / "target"),
                "run_root": str(self.fuzz_temp / "runs"),
                "tmpdir": str(self.fuzz_temp / "tmp"),
            },
            "source_binding": source_binding,
            "target_source": {
                "current": {"path": str(self.root / "current.rs"), "bytes": 1, "sha256": "c" * 64},
                "after": {"path": str(self.root / "after.rs"), "bytes": 1, "sha256": "c" * 64},
                "copied": _meta(package / "source_backed_tail_append_stream.rs"),
            },
            "manifest": {"package": _meta(package / "Cargo.toml")},
            "lock": {"package": _meta(package / "Cargo.lock")},
            "corpus_start": corpus_meta,
            "retained_0485": retained_old,
            "evidence_paths": {
                "target_source": str(evidence / "target-source.rs"),
                "manifest": str(evidence / "Cargo.toml"),
                "lock": str(evidence / "Cargo.lock"),
            },
            "evidence_files": {},
        }
        for key, source, name in (("target_source", package / "source_backed_tail_append_stream.rs", "target-source.rs"),
                                  ("manifest", package / "Cargo.toml", "Cargo.toml"),
                                  ("lock", package / "Cargo.lock", "Cargo.lock")):
            destination = evidence / name
            destination.write_bytes(source.read_bytes())
            prepared["evidence_files"][key] = _meta(destination)
        _write_json(evidence / "prepared.json", prepared)

        build_binary = {
            "origin": str(origin),
            "retained": str(retained),
            "bytes": retained.stat().st_size,
            "sha256": _sha(retained),
        }
        build = {
            "schema": "docx-stream-fuzz-build-0497-v1",
            "driver": _meta(self.root / "fuzz.py"),
            "source_binding_before": source_binding,
            "source_binding_after": source_binding,
            "prepared_sha256": _sha(evidence / "prepared.json"),
            "binary": build_binary,
        }
        _write_json(evidence / "build.json", build)

        runs: list[dict[str, object]] = []
        for seed in cleanup.FUZZ_RUN_SEEDS:
            label = f"run-{seed}"
            run_evidence = evidence / label
            run_evidence.mkdir()
            for directory in (run_evidence / "starting-corpus", run_evidence / "post-corpus",
                              run_evidence / "artifacts"):
                directory.mkdir()
            (run_evidence / "starting-corpus" / "seed.bin").write_bytes(b"seed\n")
            (run_evidence / "post-corpus" / "seed.bin").write_bytes(b"seed\n")
            (run_evidence / "terminal.stdout").write_bytes(b"")
            (run_evidence / "terminal.stderr").write_bytes(b"")
            starting_file = _meta(run_evidence / "starting-corpus" / "seed.bin")
            post_file = _meta(run_evidence / "post-corpus" / "seed.bin")
            starting = {"seed.bin": {"bytes": starting_file["bytes"],
                                     "sha256": starting_file["sha256"]}}
            post = {"seed.bin": {"bytes": post_file["bytes"],
                                  "sha256": post_file["sha256"]}}
            artifacts: dict[str, object] = {}
            run = {
                "schema": "docx-stream-fuzz-run-0497-v1",
                "label": label,
                "seed": seed,
                "exit_code": 0,
                "source_binding_before": source_binding,
                "source_binding_after": source_binding,
                "binary": {"path": str(retained), "bytes": retained.stat().st_size,
                           "sha256": _sha(retained)},
                "terminal": {
                    "stdout": str(run_evidence / "terminal.stdout"),
                    "stderr": str(run_evidence / "terminal.stderr"),
                    "stdout_meta": _meta(run_evidence / "terminal.stdout"),
                    "stderr_meta": _meta(run_evidence / "terminal.stderr"),
                },
                "corpus_before": starting,
                "corpus_after": post,
                "retained": {
                    "starting_corpus": str(run_evidence / "starting-corpus"),
                    "post_corpus": str(run_evidence / "post-corpus"),
                    "artifacts": str(run_evidence / "artifacts"),
                    "post_corpus_inventory": post,
                    "artifacts_inventory": artifacts,
                },
            }
            _write_json(run_evidence / "run.json", run)
            runs.append({"label": label, "seed": seed, "exit_code": 0})
        _write_json(evidence / "verify.json", {
            "schema": "docx-stream-fuzz-verify-0497-v1",
            "passed": True,
            "prepared_sha256": _sha(evidence / "prepared.json"),
            "build_sha256": _sha(evidence / "build.json"),
            "source_binding": source_binding,
            "run_seeds": list(cleanup.FUZZ_RUN_SEEDS),
            "binary": {"path": str(retained), "bytes": retained.stat().st_size,
                       "sha256": _sha(retained)},
            "runs": runs,
        })

    def _plan(self) -> dict[str, object]:
        return cleanup.plan_cleanup(root=self.root, temp=self.temp, target=self.target,
                                    fuzz_temp=self.fuzz_temp,
                                    build_receipts=self.receipts, proc_root=self.proc,
                                    self_pid=999999)

    def _mark_early_target_cleanup(self) -> None:
        self.target.rmdir()
        _write_json(self.root / cleanup.EARLY_TARGET_CLEANUP, {
            "path": str(self.target.resolve()),
            "reason": "fixture target was removed before final cleanup",
            "allocated_bytes": 4096,
            "free_before": 1,
            "free_after": 4097,
            "completed_ns": 1,
            "active_build_refs": [],
        })
        shutil.rmtree(self.fuzz_temp / "target")
        _write_json(self.root / cleanup.EARLY_FUZZ_CLEANUP, {
            "path": str((self.fuzz_temp / "target").resolve()),
            "allocated_bytes": 4096,
            "completed_ns": 1,
            "reason": "fixture fuzz target was removed before final cleanup",
        })

    def _fake_process(self, pid: int, *, cwd: Path | None = None,
                      exe: Path | None = None, fd: Path | None = None,
                      command: list[str] | None = None) -> None:
        process = self.proc / str(pid)
        (process / "fd").mkdir(parents=True)
        if cwd is not None:
            (process / "cwd").symlink_to(cwd, target_is_directory=True)
        if exe is not None:
            (process / "exe").symlink_to(exe)
        if fd is not None:
            (process / "fd" / "3").symlink_to(fd)
        (process / "comm").write_text("fixture\n", encoding="utf-8")
        (process / "cmdline").write_bytes(
            b"\0".join(item.encode() for item in (command or ["worker"])) + b"\0"
        )

    def test_dry_plan_is_read_only_and_reports_allocated_bytes(self) -> None:
        (self.temp / "tmp").mkdir()
        (self.temp / "tmp" / "object").write_bytes(b"scratch")
        (self.temp / "projections").mkdir()
        (self.temp / "runs").mkdir()
        (self.temp / "publication-profiles").mkdir()
        plan = self._plan()
        self.assertTrue((self.temp / "tmp" / "object").exists())
        self.assertTrue(self.target.exists())
        self.assertTrue(all(isinstance(item["allocated_bytes"], int)
                            for item in plan["removed_stats"]))

    def test_precleaned_targets_require_receipts_and_remain_read_only(self) -> None:
        self._mark_early_target_cleanup()
        plan = self._plan()
        self.assertFalse(plan["target_present"])
        self.assertFalse(plan["fuzz_target_present"])
        self.assertNotIn(self.target, plan["candidates"])
        self.assertTrue(self.fuzz_temp in plan["candidates"])
        self.assertFalse(self.target.exists())
        self.assertFalse((self.fuzz_temp / "target").exists())

        receipt = cleanup.execute_cleanup(plan)
        self.assertEqual(receipt["target_cleanup"]["mode"], "early")
        self.assertEqual(receipt["fuzz_target_cleanup"]["mode"], "early")
        proof = cleanup.verify(root=self.root, temp=self.temp, target=self.target,
                               fuzz_temp=self.fuzz_temp,
                               build_receipts=self.receipts, proc_root=self.proc)
        self.assertEqual(proof["status"], "pass")

    def test_execute_and_canonical_verify_keep_exact_retained_layout(self) -> None:
        for name in ("tmp", "projections", "runs", "publication-profiles"):
            (self.temp / name).mkdir()
        for name in ("strace1", "strace2"):
            (self.temp / "publication-profiles" / name).mkdir()
        (self.temp / "tmp" / "scratch").write_bytes(b"scratch")
        receipt = cleanup.execute_cleanup(self._plan())
        self.assertEqual(receipt["status"], "pass")
        self.assertEqual({item.name for item in self.temp.iterdir()}, {"retained"})
        self.assertFalse(self.target.exists())
        proof = cleanup.verify(root=self.root, temp=self.temp, target=self.target,
                               fuzz_temp=self.fuzz_temp,
                               build_receipts=self.receipts, proc_root=self.proc)
        self.assertEqual(proof["status"], "pass")
        self.assertEqual(proof["remaining"], [])
        self.assertEqual(
            sorted(path.relative_to(self.temp).as_posix()
                   for path in (self.temp / "retained").rglob("*") if path.is_file()),
            sorted([*(f"retained/{phase}/{role}/{cleanup.BINARY_NAME}"
                     for phase in cleanup.PHASES for role in cleanup.BUILD_ROLES),
                    f"retained/fuzz/{cleanup.FUZZ_BINARY_NAME}"]),
        )

    def test_binary_drift_aborts_before_any_removal(self) -> None:
        (self.temp / "tmp").mkdir()
        binary = self.temp / "retained" / "before" / "normal" / cleanup.BINARY_NAME
        binary.write_bytes(b"changed")
        with self.assertRaises(cleanup.CleanupError):
            self._plan()
        self.assertTrue((self.temp / "tmp").exists())
        self.assertTrue(self.target.exists())

    def test_incomplete_terminal_gate_is_refused(self) -> None:
        (self.root / "validation" / "unfinished.started.json").write_text(
            "{}\n", encoding="utf-8"
        )
        with self.assertRaisesRegex(cleanup.CleanupError, "not complete"):
            self._plan()
        self.assertTrue(self.target.exists())

    def test_symlink_and_special_candidate_are_refused(self) -> None:
        outside = Path(self.holder.name) / "outside"
        outside.mkdir()
        (outside / "secret").write_bytes(b"keep")
        (self.temp / "tmp").symlink_to(outside, target_is_directory=True)
        with self.assertRaises(cleanup.CleanupError):
            self._plan()
        self.assertTrue((outside / "secret").exists())
        (self.temp / "tmp").unlink()
        os.mkfifo(self.temp / "tmp")
        with self.assertRaises(cleanup.CleanupError):
            self._plan()

    def test_process_cwd_exe_and_fd_references_block_plan(self) -> None:
        candidate = self.temp / "tmp"
        candidate.mkdir()
        opened = candidate / "opened"
        opened.write_bytes(b"opened")
        executable = candidate / "worker"
        executable.write_bytes(b"worker")
        executable.chmod(executable.stat().st_mode | stat.S_IXUSR)
        self._fake_process(101, cwd=candidate, exe=executable, fd=opened)
        audit = cleanup._process_audit([candidate], self.root, self.temp, self.target,
                                       proc_root=self.proc, self_pid=999999)
        self.assertFalse(audit["safe"])
        self.assertEqual({item["reference"] for item in audit["candidate_references"]},
                         {"cwd", "exe", "fd:3"})

    def test_inaccessible_foreign_process_with_explicit_scratch_reference_blocks(self) -> None:
        process = self.proc / "105"
        self._fake_process(
            105,
            command=[str(self.root / "measure.py"), str(self.temp / "target")],
        )

        def inaccessible(_: Path) -> Path:
            raise cleanup.CleanupError("foreign proc link hidden")

        with patch.object(cleanup.os, "getuid", return_value=4242):
            with patch.object(cleanup, "_proc_link", side_effect=inaccessible):
                audit = cleanup._process_audit(
                    [], self.root, self.temp, self.target,
                    proc_root=self.proc, self_pid=999999,
                )
        self.assertFalse(audit["safe"])
        self.assertEqual(audit["foreign_unobservable_processes"], [])
        self.assertEqual(len(audit["root_gate_capture_processes"]), 1)
        driver = audit["root_gate_capture_processes"][0]
        self.assertEqual(driver["pid"], 105)
        self.assertTrue(driver["workspace_reference"])
        self.assertTrue(driver["script_reference"])
        self.assertNotIn("command", driver)
        self.assertGreater(driver["command_bytes"], 0)
        self.assertEqual(len(driver["command_sha256"]), 64)
        self.assertEqual(driver["unobservable_references"], ["cwd", "exe", "fd"])

    def test_inaccessible_unknown_same_uid_process_fails_closed(self) -> None:
        process = self.proc / "106"
        self._fake_process(106, command=["unrelated-worker"])

        def inaccessible(_: Path) -> Path:
            raise cleanup.CleanupError("same uid proc link hidden")

        with patch.object(cleanup, "_proc_link", side_effect=inaccessible):
            with self.assertRaisesRegex(cleanup.CleanupError, "same uid proc link hidden"):
                cleanup._process_audit(
                    [], self.root, self.temp, self.target,
                    proc_root=self.proc, self_pid=999999,
                )

    def test_live_root_capture_process_blocks_even_without_scratch_reference(self) -> None:
        (self.temp / "tmp").mkdir()
        python = Path(sys.executable)
        self._fake_process(102, cwd=self.root, exe=python,
                           command=[str(python), "measure.py"])
        with self.assertRaisesRegex(cleanup.CleanupError, "live 0497 gate/capture"):
            self._plan()
        self.assertTrue((self.temp / "tmp").exists())

    def test_live_fuzz_driver_executable_is_blocked(self) -> None:
        (self.temp / "tmp").mkdir()
        self._fake_process(104, cwd=self.root, exe=self.root / "fuzz.py",
                           command=[str(self.root / "fuzz.py"), "run"])
        with self.assertRaisesRegex(cleanup.CleanupError, "live 0497 gate/capture"):
            self._plan()
        self.assertTrue(self.fuzz_temp.exists())

    def test_failed_fuzz_run_keeps_fuzz_scratch_and_refuses_cleanup(self) -> None:
        run_path = self.root / "fuzz" / "run-498" / "run.json"
        value = json.loads(run_path.read_text(encoding="utf-8"))
        value["exit_code"] = 1
        _write_json(run_path, value)
        with self.assertRaisesRegex(cleanup.CleanupError, "terminal did not pass"):
            self._plan()
        self.assertTrue(self.fuzz_temp.exists())

    def test_fuzz_evidence_copy_drift_blocks_scratch_removal(self) -> None:
        (self.root / "fuzz" / "run-497" / "terminal.stdout").write_bytes(b"tampered")
        with self.assertRaisesRegex(cleanup.CleanupError, "content changed"):
            self._plan()
        self.assertTrue(self.fuzz_temp.exists())

    def test_failed_private_scratch_is_reported_and_not_destructively_planned(self) -> None:
        private = self.temp / "runs" / "attempt" / "label"
        private.mkdir(parents=True)
        (private / "failed.trace").write_bytes(b"retain")
        terminal = self.root / "validation" / "fixture.json"
        value = json.loads(terminal.read_text(encoding="utf-8"))
        value["cleanup"] = {
            "status": "preserved_failure",
            "root": str(private),
            "remaining": [str(private)],
        }
        _write_json(terminal, value)
        plan = self._plan()
        self.assertTrue(plan["preserved_failed"])
        with self.assertRaisesRegex(cleanup.CleanupError, "archived custody"):
            cleanup.execute_cleanup(plan)
        self.assertTrue((private / "failed.trace").exists())
        self.assertTrue(self.target.exists())

    def test_destructive_cli_rejects_fixture_roots(self) -> None:
        (self.temp / "tmp").mkdir()
        status = cleanup.main(["--root", str(self.root), "--temp-root", str(self.temp),
                               "--target-root", str(self.target), "--proc-root", str(self.proc),
                               "--fuzz-temp-root", str(self.fuzz_temp),
                               "--before-build", str(self.receipts["before"]),
                               "--after-build", str(self.receipts["after"])])
        self.assertEqual(status, 1)
        self.assertTrue((self.temp / "tmp").exists())
        self.assertFalse((self.root / "cleanup.json").exists())


if __name__ == "__main__":
    raise SystemExit(unittest.main())
