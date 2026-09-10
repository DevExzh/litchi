#!/usr/bin/env python3
"""Focused safety tests for the 0497 ENOSPC continuation driver."""

from __future__ import annotations

import importlib.util
import datetime as _datetime
import json
from pathlib import Path
import tempfile
import unittest
from unittest import mock


ROOT = Path(__file__).resolve().parent
SPEC = importlib.util.spec_from_file_location("resume_0497_tested", ROOT / "resume.py")
assert SPEC is not None and SPEC.loader is not None
RESUME = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(RESUME)


def _protocol(count: int = 288) -> dict[str, object]:
    return {"formal_runs": [{"ordinal": ordinal, "label": f"run-{ordinal}",
                              "phase": "before", "role": "normal"}
                             for ordinal in range(count)]}


class ResumeInventoryTests(unittest.TestCase):
    def test_suffix_is_exactly_the_frozen_remaining_97_runs(self) -> None:
        suffix = RESUME.remaining_specs(_protocol())
        self.assertEqual(len(suffix), 97)
        self.assertEqual(suffix[0], {"ordinal": 191, "label": "run-191",
                                     "phase": "before", "role": "normal"})
        self.assertEqual(suffix[-1], {"ordinal": 287, "label": "run-287",
                                      "phase": "before", "role": "normal"})

    def test_suffix_rejects_ordinal_gap(self) -> None:
        protocol = _protocol()
        protocol["formal_runs"][191]["ordinal"] = 192  # type: ignore[index]
        with self.assertRaisesRegex(RESUME.ResumeError, "ordinal"):
            RESUME.remaining_specs(protocol)

    def test_suffix_rejects_reordered_or_short_inventory(self) -> None:
        protocol = _protocol()
        protocol["formal_runs"][190], protocol["formal_runs"][191] = (
            protocol["formal_runs"][191], protocol["formal_runs"][190],  # type: ignore[index]
        )
        with self.assertRaises(RESUME.ResumeError):
            RESUME.remaining_specs(protocol)
        with self.assertRaises(RESUME.ResumeError):
            RESUME.remaining_specs(_protocol(287))

    def test_only_fixed_start_ordinal_is_supported(self) -> None:
        with self.assertRaisesRegex(RESUME.ResumeError, "only supports ordinal 191"):
            RESUME.remaining_specs(_protocol(), start=190)

    def test_launch_specs_bind_the_attempt_and_exact_suffix_order(self) -> None:
        specs = RESUME.launch_specs(_protocol())
        self.assertEqual([item["ordinal"] for item in specs], list(range(191, 288)))
        self.assertTrue(all(item["attempt"] == RESUME.ATTEMPT for item in specs))

    def test_launch_remaining_records_completed_prefix_before_failure(self) -> None:
        protocol = _protocol()
        specs = RESUME.launch_specs(protocol)
        completed: list[str] = []

        class FakeMeasure:
            DEFAULT_TIMEOUT_SECONDS = 1800

            def __init__(self) -> None:
                self.calls: list[dict[str, object]] = []

            def _build_key(self, phase: str, role: str) -> str:
                return f"{phase}/{role}"

            def _launch(self, spec: dict[str, object], build: object,
                        frozen: object, protocol_hash: str, timeout: int) -> None:
                self.calls.append(spec)
                if spec["ordinal"] == 193:
                    raise RuntimeError("synthetic child failure")

        measure = FakeMeasure()
        builds = {"before/normal": {}}
        with self.assertRaisesRegex(RuntimeError, "synthetic child failure"):
            RESUME.launch_remaining(measure, protocol, "protocol-hash", builds, specs,
                                    completed=completed)
        self.assertEqual([item["ordinal"] for item in measure.calls], [191, 192, 193])
        self.assertEqual(completed, ["run-191", "run-192"])
        self.assertTrue(all(item["attempt"] == RESUME.ATTEMPT for item in measure.calls))


class InterruptedCustodyTests(unittest.TestCase):
    def test_raw_archive_names_cannot_be_seen_as_live_receipts(self) -> None:
        for name in RESUME.REQUIRED_RAW_FILES:
            archived = RESUME.raw_archive_name(name)
            self.assertTrue(archived.endswith(".raw"))
            self.assertNotEqual(archived, name)
            self.assertNotIn(archived, ("started.json", "terminal.json"))

    def test_raw_archive_rejects_unexpected_names(self) -> None:
        with self.assertRaises(RESUME.ResumeError):
            RESUME.raw_archive_name("terminal.json")
        with self.assertRaises(RESUME.ResumeError):
            RESUME.raw_archive_name("nested/started.json")

    def test_tree_inventory_keeps_empty_private_tmp_directory(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary) / "private"
            (root / "tmp").mkdir(parents=True)
            inventory = RESUME._tree_inventory(root)
            self.assertEqual(inventory, [{
                "path": str(root / "tmp"), "relative": "tmp", "kind": "directory",
            }])

    def test_write_new_allows_multiple_immutable_receipts_in_one_parent(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary) / "resume" / "formal1"
            first = root / "resume-start.json"
            second = root / "resume-terminal.json"
            RESUME._write_new(first, {"status": "running"})
            RESUME._write_new(second, {"status": "pass"})
            self.assertEqual(first.read_text(encoding="utf-8").count("running"), 1)
            with self.assertRaisesRegex(RESUME.ResumeError, "replace existing"):
                RESUME._write_new(first, {"status": "changed"})

    def test_resume_terminal_receipts_bind_real_time_and_exit_status(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            start_path = root / "resume-start.json"
            start_path.write_text(json.dumps({"status": "running"}) + "\n", encoding="utf-8")
            started = _datetime.datetime.now(_datetime.timezone.utc).isoformat().replace(
                "+00:00", "Z")
            start = {"remaining_children": 97, "started_utc": started}
            with mock.patch.object(RESUME, "RESUME_ROOT", root):
                for status, expected_exit in (("pass", 0), ("failed", 1)):
                    terminal = RESUME._resume_terminal(start, status, 2)
                    self.assertEqual(terminal["exit_code"], expected_exit)
                    self.assertEqual(terminal["started_utc"], started)
                    parsed = _datetime.datetime.fromisoformat(
                        terminal["finished_utc"].replace("Z", "+00:00"))
                    self.assertIsNotNone(parsed.tzinfo)


if __name__ == "__main__":
    unittest.main()
