#!/usr/bin/env python3
"""Small, source-only controls for the Q472a suffix diagnostic policy."""
from __future__ import annotations

import importlib.util
from pathlib import Path
import unittest


SCRIPT = Path(__file__).with_name("run_q472a_rest29.py")
SPEC = importlib.util.spec_from_file_location("q472a_rest29", SCRIPT)
assert SPEC is not None and SPEC.loader is not None
MODULE = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(MODULE)


class Rest29PolicyTests(unittest.TestCase):
    def test_ordinary_code_failure_continues(self) -> None:
        row = {"status": "failed", "effective_exit": 1}
        self.assertEqual(MODULE.classify_row(row), "code")
        self.assertEqual(MODULE.continuation_action(row), "continue")

    def test_resource_failure_stops(self) -> None:
        row = {"status": "failed", "resource_failure": "target hard cap"}
        self.assertEqual(MODULE.classify_row(row), "resource")
        self.assertEqual(MODULE.continuation_action(row), "stop")

    def test_provenance_and_leftover_failures_stop(self) -> None:
        for row, expected in (
            ({"status": "failed", "provenance_failure": "source changed"}, "provenance"),
            ({"status": "failed", "owned_processes": [{"pid": 7}]}, "leftover-process"),
        ):
            self.assertEqual(MODULE.classify_row(row), expected)
            self.assertEqual(MODULE.continuation_action(row), "stop")

    def test_officecli_preservation_failure_stops(self) -> None:
        row = {"status": "failed", "barrier_failure": "TSV missing"}
        self.assertEqual(MODULE.classify_row(row), "preservation")
        self.assertEqual(MODULE.continuation_action(row), "stop")

    def test_identity_drift_is_fail_closed(self) -> None:
        for label in ("source", "plan", "tools"):
            with self.assertRaises(MODULE.DiagnosticError):
                MODULE.require_identity(label, "pinned", "drifted")

    def test_exact_suffix_and_skipped_prefix(self) -> None:
        self.assertEqual(MODULE.FULL53_SUFFIX, tuple(range(29, 53)))
        self.assertEqual(MODULE.RUNTIME_SUFFIX, tuple(range(7, 12)))
        self.assertEqual(len(MODULE.FULL53_SUFFIX) + len(MODULE.RUNTIME_SUFFIX), 29)
        self.assertEqual(len(range(29)), 29)


if __name__ == "__main__":
    unittest.main()
