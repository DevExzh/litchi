"""Source-only tests for the PPT build receipt projection adapter."""
from __future__ import annotations

import copy
import importlib.util
from pathlib import Path
import tempfile
import unittest


ROOT = Path(__file__).resolve().parents[1]
SPEC = importlib.util.spec_from_file_location(
    "ppt_lazy_build_receipt_adapter",
    ROOT / "tools/ppt_lazy_build_receipt_adapter.py",
)
assert SPEC and SPEC.loader
MODULE = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(MODULE)


class BuildReceiptAdapterTests(unittest.TestCase):
    def test_projection_removes_only_receipt_content_identity(self) -> None:
        manifest = {
            "build": {
                "receipt": {
                    "path": "/outside/receipt.json",
                    "bytes": 123,
                    "sha256": "a" * 64,
                },
                "profile": "release",
            },
            "protocol": {"cpu": 29},
        }
        projection = MODULE.projection_value(manifest)
        self.assertEqual(
            projection,
            {
                "build": {
                    "receipt": {"path": "/outside/receipt.json"},
                    "profile": "release",
                },
                "protocol": {"cpu": 29},
            },
        )
        self.assertEqual(manifest["build"]["receipt"]["bytes"], 123)
        self.assertEqual(manifest["build"]["receipt"]["sha256"], "a" * 64)

    def test_projection_rejects_extra_receipt_fields(self) -> None:
        manifest = {
            "build": {
                "receipt": {
                    "path": "/outside/receipt.json",
                    "bytes": 123,
                    "sha256": "a" * 64,
                    "status": "passed",
                }
            }
        }
        with self.assertRaises(MODULE.AdapterError):
            MODULE.projection_value(manifest)

    def test_projection_is_sensitive_to_every_non_receipt_input(self) -> None:
        manifest = {
            "build": {
                "receipt": {"path": "/outside/receipt.json", "bytes": 123, "sha256": "a" * 64},
                "commands": {"normal": ["cargo", "build"]},
            },
            "candidate": {"ref": "candidate"},
        }
        expected = MODULE.canonical(MODULE.projection_value(manifest))
        changed = copy.deepcopy(manifest)
        changed["candidate"]["ref"] = "different"
        self.assertNotEqual(expected, MODULE.canonical(MODULE.projection_value(changed)))

    def test_strict_json_rejects_nonfinite_and_duplicate_values(self) -> None:
        with tempfile.TemporaryDirectory(prefix="ppt-build-adapter-") as directory:
            path = Path(directory) / "input.json"
            path.write_text('{"value": 1e999}', encoding="utf-8")
            with self.assertRaises(MODULE.AdapterError):
                MODULE.strict_json(path, "test")
            path.write_text('{"value": 1, "value": 2}', encoding="utf-8")
            with self.assertRaises(MODULE.AdapterError):
                MODULE.strict_json(path, "test")


if __name__ == "__main__":
    unittest.main()
