#!/usr/bin/env python3
"""Fixture-only tests for the 0497 immutable evidence seal."""

from __future__ import annotations

import json
import os
from pathlib import Path
import tempfile
import unittest
from types import SimpleNamespace
from unittest.mock import patch


import seal


TEST_TMP = Path(os.environ.get("TMPDIR", tempfile.gettempdir())) / "litchi-goal-0497-test-tmp"
TEST_TMP.mkdir(parents=True, exist_ok=True)


class SealFixtureTests(unittest.TestCase):
    def setUp(self) -> None:
        self.holder = tempfile.TemporaryDirectory(prefix="litchi-seal-0497-", dir=TEST_TMP)
        self.root = Path(self.holder.name) / "evidence"
        self.root.mkdir()

    def tearDown(self) -> None:
        self.holder.cleanup()

    def test_completed_resume_reuses_capture_path_and_releases_private_root(self) -> None:
        capture = self.root / "completed-capture"
        private = self.root / "private"
        capture.mkdir()
        seal._resume_live_roots(capture, private)
        private.mkdir()
        with self.assertRaises(seal.SealError):
            seal._resume_live_roots(capture, private)
        private.rmdir()
        capture.rmdir()
        with self.assertRaises(seal.SealError):
            seal._resume_live_roots(capture, private)
        capture.symlink_to(self.root, target_is_directory=True)
        with self.assertRaises(seal.SealError):
            seal._resume_live_roots(capture, private)

    def test_inventory_excludes_only_seal_and_transient_bytecode(self) -> None:
        (self.root / "README.md").write_text("evidence\n", encoding="utf-8")
        (self.root / "seal.json").write_text("future\n", encoding="utf-8")
        transient = self.root / "nested" / "__pycache__"
        transient.mkdir(parents=True)
        (transient / "helper.pyc").write_bytes(b"bytecode")
        self.assertEqual(set(seal._inventory(self.root)), {"README.md"})

    def test_inventory_rejects_links_bytecode_and_special_files(self) -> None:
        outside = Path(self.holder.name) / "outside"
        outside.write_text("outside\n", encoding="utf-8")
        (self.root / "link").symlink_to(outside)
        with self.assertRaises(seal.SealError):
            seal._inventory(self.root)

        (self.root / "link").unlink()
        (self.root / "unexpected.pyc").write_bytes(b"bytecode")
        with self.assertRaises(seal.SealError):
            seal._inventory(self.root)
        (self.root / "unexpected.pyc").unlink()

        if not hasattr(os, "mkfifo"):
            self.skipTest("FIFO fixtures are unavailable")
        fifo = self.root / "fifo"
        os.mkfifo(fifo)
        with self.assertRaises(seal.SealError):
            seal._inventory(self.root)

    def test_descriptors_reject_external_and_symlink_component_paths(self) -> None:
        outside = Path(self.holder.name) / "outside.txt"
        outside.write_text("outside\n", encoding="utf-8")
        with self.assertRaises(seal.SealError):
            seal._descriptor(self.root, outside, "outside")

        linked = self.root / "linked"
        linked.symlink_to(Path(self.holder.name), target_is_directory=True)
        with self.assertRaises(seal.SealError):
            seal._descriptor(self.root, linked / "outside.txt", "linked")

    def test_capture_inventory_authenticates_each_raw_artifact(self) -> None:
        captures = self.root / "captures"
        first = captures / "first"
        second = captures / "second"
        first.mkdir(parents=True)
        second.mkdir()
        (first / "report.json").write_text("{}\n", encoding="utf-8")
        (first / "stdout.txt").write_text("ok\n", encoding="utf-8")
        (second / "report.json").write_text("{}\n", encoding="utf-8")
        binding = seal._capture_files(
            self.root,
            [
                {"directory": first, "spec": {"label": "z-case"}},
                {"directory": second, "spec": {"label": "a-case"}},
            ],
            "formal1",
        )
        self.assertEqual(binding["children"], 2)
        self.assertEqual([item["label"] for item in binding["captures"]], ["a-case", "z-case"])
        self.assertEqual(binding["captures"][1]["files"]["report.json"]["path"],
                         "captures/first/report.json")

    def test_capture_inventory_rejects_special_artifacts(self) -> None:
        capture = self.root / "capture"
        capture.mkdir()
        if not hasattr(os, "mkfifo"):
            self.skipTest("FIFO fixtures are unavailable")
        os.mkfifo(capture / "report.json")
        with self.assertRaises(seal.SealError):
            seal._capture_files(self.root, [{"directory": capture, "spec": {"label": "case"}}], "formal1")

    def test_source_binding_includes_verified_fixture_phases(self) -> None:
        source = Path(self.holder.name) / "repo" / "src.rs"
        source.parent.mkdir()
        source.write_text("fn fixture() {}\n", encoding="utf-8")
        (self.root / "candidate-source.json").write_text("{}\n", encoding="utf-8")
        (self.root / "candidate.patch").write_text("patch\n", encoding="utf-8")
        fixture = self.root / "fixture-inputs.json"
        fixture.write_text(json.dumps({"scope": "fixture", "files": {"src.rs": {
            "bytes": source.stat().st_size, "sha256": seal._sha(source)}}}) + "\n",
                           encoding="utf-8")
        fixture_verification = self.root / "fixture-verification.json"
        with patch.object(seal, "TEMP", Path(self.holder.name) / "scratch"):
            fixture_verification.write_text(json.dumps({
                "verified_utc": seal._now(),
                "manifest_sha256": seal._sha(fixture),
                "phases": {
                    phase: {"source_root": str((seal.TEMP / phase).resolve(strict=False)),
                            "checked_files": 1, "checked_bytes": 1, "status": "pass"}
                    for phase in ("before", "after")
                },
            }) + "\n", encoding="utf-8")
            class FakeMeasure:
                @staticmethod
                def _candidate_source_binding():
                    return ({"manifest": {}, "patch": {}}, {
                        "src.rs": {"bytes": source.stat().st_size, "sha256": seal._sha(source)},
                    })

                @staticmethod
                def _fixture_inputs_binding():
                    return seal._meta(fixture, "fixture inputs")

            binding = seal._source_binding(FakeMeasure, self.root, source.parent)
        self.assertEqual(binding["fixture"]["verification"]["phases"]["after"]["status"], "pass")
        self.assertEqual(binding["fixture"]["verification"]["manifest_sha256"], seal._sha(fixture))

    def test_projection_cleanup_is_empty_directory_only(self) -> None:
        capture = SimpleNamespace(TEMP=Path(self.holder.name) / "scratch")
        projection = capture.TEMP / "projections"
        projection.mkdir(parents=True)
        seal._remove_empty(projection, "projection scratch")
        self.assertFalse(projection.exists())

        projection.mkdir(parents=True)
        marker = projection / "unexpected"
        marker.write_text("keep\n", encoding="utf-8")
        with self.assertRaises(seal.SealError):
            seal._remove_empty(projection, "projection scratch")
        self.assertTrue(marker.exists())

    def test_early_target_cleanup_binds_absent_target_custody(self) -> None:
        target = Path(self.holder.name) / "target"
        receipt = self.root / seal.EARLY_TARGET_CLEANUP
        receipt.write_text(json.dumps({
            "path": str(target),
            "reason": "fixture target was removed after terminal builds",
            "allocated_bytes": 123,
            "free_before": 10,
            "free_after": 133,
            "completed_ns": 1,
            "active_build_refs": [],
        }) + "\n", encoding="utf-8")
        with patch.object(seal, "TARGET", target):
            binding = seal._early_target_binding(self.root)
        self.assertEqual(binding["path"], str(target))
        self.assertEqual(binding["allocated_bytes"], 123)
        self.assertEqual(binding["active_build_refs"], [])

        value = json.loads(receipt.read_text(encoding="utf-8"))
        value["active_build_refs"] = ["active-build"]
        receipt.write_text(json.dumps(value) + "\n", encoding="utf-8")
        with patch.object(seal, "TARGET", target):
            with self.assertRaises(seal.SealError):
                seal._early_target_binding(self.root)

    def test_resume_empty_private_archive_may_be_omitted_but_files_are_rejected(self) -> None:
        private = self.root / "interrupted" / "private-root"
        declared = [{"path": str(private / "tmp"), "relative": "tmp", "kind": "directory"}]
        self.assertEqual(seal._resume_optional_directory_inventory(private, declared, "fixture"), [])
        (private / "tmp").mkdir(parents=True)
        self.assertEqual(seal._resume_optional_directory_inventory(private, declared, "fixture"), [
            {"path": str(private / "tmp"), "relative": "tmp", "kind": "directory"},
        ])
        (private / "tmp" / "unexpected.txt").write_text("unexpected\n", encoding="utf-8")
        with self.assertRaises(seal.SealError):
            seal._resume_optional_directory_inventory(private, declared, "fixture")

    def test_timestamps_are_timezone_bound_and_utc(self) -> None:
        current = seal._now()
        seal._timestamp(current, "current")
        with self.assertRaises(seal.SealError):
            seal._timestamp("2026-09-10T12:00:00", "naive")
        with self.assertRaises(seal.SealError):
            seal._timestamp("2026-09-10T12:00:00+01:00", "non-utc")
        with self.assertRaises(seal.SealError):
            seal._timestamp("not-a-timestamp", "malformed")

    def test_manifest_is_exclusive_and_verifies_immutable_inventory(self) -> None:
        (self.root / "README.md").write_text("evidence\n", encoding="utf-8")
        evidence = {
            "protocol": {}, "builds": {}, "formal": {}, "pilot": {},
            "analysis": {}, "verification": {}, "resume": {}, "profiles": {}, "fuzz": {},
            "early_target_cleanup": {}, "early_fuzz_target_cleanup": {}, "cleanup": {},
            "source": {}, "adr": {}, "reviews": [],
            "gates": {}, "protected_primary": {},
        }
        with patch.object(seal, "_evidence", return_value=evidence):
            manifest = seal._manifest(self.root)
        self.assertEqual(manifest["files"]["README.md"]["bytes"], len("evidence\n"))
        seal_path = self.root / "seal.json"
        seal_path.write_text(json.dumps(manifest) + "\n", encoding="utf-8")
        with patch.object(seal, "_evidence", return_value=evidence):
            seal._verify_manifest(self.root, manifest)
            with self.assertRaises(seal.SealError):
                seal._manifest(self.root)
        (self.root / "README.md").write_text("changed\n", encoding="utf-8")
        with patch.object(seal, "_evidence", return_value=evidence):
            with self.assertRaises(seal.SealError):
                seal._verify_manifest(self.root, manifest)

    def _make_gate_fixture(self, *, include_helpers: bool = False) -> dict[str, list[str]]:
        driver = self.root / "gate.py"
        driver.write_text("gate fixture\n", encoding="utf-8")
        validation = self.root / "validation"
        validation.mkdir()
        commands: dict[str, list[str]] = {}
        receipts: dict[str, dict[str, int | str]] = {}
        names = set(seal.MANDATORY_GATE_NAMES)
        if include_helpers:
            names.update(seal.OPTIONAL_GATE_NAMES)
        for index, name in enumerate(sorted(names)):
            command = ["python3", "-B", name]
            commands[name] = command
            source = validation / f"{name}.source.json"
            stdout = validation / f"{name}.stdout"
            stderr = validation / f"{name}.stderr"
            started_path = validation / f"{name}.started.json"
            receipt_path = validation / f"{name}.json"
            source.write_text(json.dumps({"src.rs": {"path": "/fixture/src.rs",
                                                       "bytes": 1, "sha256": "0" * 64}}) + "\n",
                                encoding="utf-8")
            stdout.write_text(f"stdout {name}\n", encoding="utf-8")
            stderr.write_text("", encoding="utf-8")
            metadata = {
                "argv": command,
                "cwd": str((seal.TEMP / "after").resolve(strict=False)),
                "environment": dict(seal.GATE_ENVIRONMENT),
                "driver": seal._meta(driver),
                "source_manifest": seal._meta(source),
                "started_ns": index + 1,
                "timeout_seconds": 3600,
            }
            started_path.write_text(json.dumps(metadata) + "\n", encoding="utf-8")
            receipt = dict(metadata)
            receipt.update({
                "pid": 1000 + index,
                "timed_out": False,
                "termination": None,
                "exit_code": 0,
                "finished_ns": index + 2,
                "source_unchanged": True,
                "stdout": seal._meta(stdout),
                "stderr": seal._meta(stderr),
            })
            receipt_path.write_text(json.dumps(receipt) + "\n", encoding="utf-8")
            receipts[name] = {"path": f"validation/{name}.json",
                              "bytes": receipt_path.stat().st_size,
                              "sha256": seal._sha(receipt_path)}
        (self.root / "final-gates.json").write_text(
            json.dumps({"schema": "final-gates-v1", "version": 1,
                        "scope": {"required": True}, "commands": commands,
                        "receipts": receipts}) + "\n",
            encoding="utf-8",
        )
        (self.root / "gates-candidate8.json").write_text("{}\n", encoding="utf-8")
        return commands

    def test_final_gate_binding_records_exact_commands_and_receipts(self) -> None:
        commands = self._make_gate_fixture(include_helpers=True)
        binding = seal._gates_binding(self.root)
        self.assertEqual(set(binding["final"]["commands"]), set(commands))
        self.assertEqual(set(binding["final"]["receipts"]), set(commands))
        self.assertIn("helpers", binding["final"]["receipts"])
        self.assertEqual(binding["final"]["receipts"]["focused"]["exit_code"], 0)
        self.assertEqual(len(binding["historical"]), 1)

    def test_final_gate_binding_rejects_terminal_failure(self) -> None:
        self._make_gate_fixture()
        path = self.root / "validation/focused.json"
        receipt = json.loads(path.read_text(encoding="utf-8"))
        receipt["exit_code"] = 1
        path.write_text(json.dumps(receipt) + "\n", encoding="utf-8")
        with self.assertRaises(seal.SealError):
            seal._gates_binding(self.root)

    def test_final_gate_binding_rejects_command_drift(self) -> None:
        commands = self._make_gate_fixture()
        commands["focused"] = ["python3", "-B", "changed"]
        value = json.loads((self.root / "final-gates.json").read_text(encoding="utf-8"))
        value["commands"] = commands
        (self.root / "final-gates.json").write_text(json.dumps(value) + "\n", encoding="utf-8")
        with self.assertRaises(seal.SealError):
            seal._gates_binding(self.root)

    def test_final_gate_binding_rejects_stdout_metadata_drift(self) -> None:
        self._make_gate_fixture()
        stdout = self.root / "validation/focused.stdout"
        stdout.write_text("changed\n", encoding="utf-8")
        with self.assertRaises(seal.SealError):
            seal._gates_binding(self.root)

    def test_final_gate_binding_rejects_source_metadata_drift(self) -> None:
        self._make_gate_fixture()
        source = self.root / "validation/focused.source.json"
        source.write_text("changed\n", encoding="utf-8")
        with self.assertRaises(seal.SealError):
            seal._gates_binding(self.root)

    def test_profile_helper_is_allowlisted_and_hash_bound(self) -> None:
        helper = self.root / "profile.py"
        helper.write_text("profile helper\n", encoding="utf-8")
        metadata = seal._meta(helper)
        path, binding = seal._profile_helper_binding(self.root, metadata, "profile helper")
        self.assertEqual(path, helper)
        self.assertEqual(binding["path"], "profile.py")

        helper.write_text("changed helper\n", encoding="utf-8")
        with self.assertRaises(seal.SealError):
            seal._profile_helper_binding(self.root, metadata, "profile helper")

        unknown = self.root / "profile_other.py"
        unknown.write_text("other helper\n", encoding="utf-8")
        with self.assertRaises(seal.SealError):
            seal._profile_helper_binding(self.root, seal._meta(unknown), "profile helper")

    def test_failed_strace1_cannot_be_silently_marked_successful(self) -> None:
        helper = self.root / "profile.py"
        helper.write_text("profile helper\n", encoding="utf-8")
        attempt = self.root / "profiles" / "strace1"
        attempt.mkdir(parents=True)
        result = {
            "schema": "docx-tail-append-publication-profile-result-v1",
            "version": 1,
            "attempt": "strace1",
            "status": "pass",
            "modes": ["counting", "atomic"],
            "helper": seal._meta(helper),
        }
        (attempt / "result.json").write_text(json.dumps(result) + "\n", encoding="utf-8")
        with self.assertRaisesRegex(seal.SealError, "silently marked successful"):
            seal._profile_binding(self.root)


if __name__ == "__main__":
    unittest.main()
