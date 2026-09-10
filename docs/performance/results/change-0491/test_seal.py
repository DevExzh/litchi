#!/usr/bin/env python3
"""Focused custody checks for the 0491 evidence seal."""

from __future__ import annotations

import json
from pathlib import Path
import sys
import tempfile
import unittest
from unittest import mock


HERE = Path(__file__).resolve().parent
if str(HERE) not in sys.path:
    sys.path.insert(0, str(HERE))

import cleanup  # noqa: E402
import seal  # noqa: E402


class SealCustodyTests(unittest.TestCase):
    def test_build_selection_must_use_canonical_receipt_path(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            (root / "alternate-build.json").write_text("{}\n", encoding="utf-8")
            builds = {
                "normal": {"path": str(root / "build-normal.json")},
                "allocator": {"path": str(root / "build-allocator.json")},
            }
            with mock.patch.object(seal, "_strict_builds", return_value=builds):
                with self.assertRaises(seal.SealError):
                    seal._build(root, "alternate-build.json", "normal", builds)

    def test_required_final_gates_are_a_subset_and_extra_gates_are_allowed(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            labels = sorted(seal.REQUIRED_FINAL_GATE_LABELS | {"boundary-final5"})

            def fake_gate(_root: Path, value: str | Path) -> dict[str, str]:
                path = Path(value)
                return {"path": str(path), "label": path.stem}

            with mock.patch.object(seal, "_gate", side_effect=fake_gate):
                selected = seal._gates(root, [f"validation/{label}.json" for label in labels])
                self.assertEqual(len(selected), len(labels))
                with self.assertRaises(seal.SealError):
                    seal._gates(root, [f"validation/{label}.json" for label in labels[:-1]])

    def test_cleanup_delegates_to_canonical_read_only_verifier(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            path = root / "cleanup.json"
            path.write_text(
                json.dumps({
                    "schema": seal.CLEANUP_SCHEMA,
                    "version": 1,
                    "status": "pass",
                }) + "\n",
                encoding="utf-8",
            )
            descriptor = cleanup._descriptor(path, "cleanup receipt cleanup.json", root=root)
            checked = {
                "schema": seal.CLEANUP_VERIFICATION_SCHEMA,
                "version": 1,
                "status": "pass",
                "cleanup_receipt": descriptor,
                "build_target_removed": True,
                "remaining": [],
            }
            with mock.patch.object(cleanup, "verify", return_value=checked) as verifier:
                result = seal._cleanup(root, "cleanup.json")
            verifier.assert_called_once_with(root=root)
            self.assertEqual(result["schema"], seal.CLEANUP_SCHEMA)
            self.assertEqual(result["verification_schema"], seal.CLEANUP_VERIFICATION_SCHEMA)
            self.assertEqual(result["remaining"], [])


if __name__ == "__main__":
    unittest.main()
