#!/usr/bin/env python3
"""Static checks for the external DrawingML compile preparer."""

from __future__ import annotations

from pathlib import Path
import sys
import tempfile
import unittest


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))

from native_qualification import drawingml_compile_driver as driver  # noqa: E402


class DrawingMlCompileDriverTests(unittest.TestCase):
    def test_harness_manifest_has_explicit_local_closure_and_pinned_hash(self) -> None:
        with tempfile.TemporaryDirectory(prefix="drawingml-compile-driver-") as directory:
            result = driver._write_harness(
                Path(directory) / "work",
                Path("/external/drawingml-source"),
                ROOT,
            )
            manifest = (Path(directory) / "work" / "Cargo.toml").read_text(encoding="utf-8")

        self.assertEqual(result["probe"]["bytes"], (ROOT / driver.PROBE_RELATIVE_PATH).stat().st_size)
        self.assertIn(
            'litchi-drawingml = { path = "/external/drawingml-source/crates/litchi-drawingml" }',
            manifest,
        )
        self.assertIn(
            'litchi-ooxml-common = { path = "/external/drawingml-source/crates/litchi-ooxml-common" }',
            manifest,
        )
        self.assertIn('sha2 = "=0.11.0"', manifest)
        for feature in driver.FEATURES:
            self.assertIn(f"{feature} = []", manifest)

    def test_revision_validation_is_fail_closed(self) -> None:
        with self.assertRaises(driver.DriverError):
            driver._require_full_sha("short", "revision")

    def test_driver_identity_is_source_bound(self) -> None:
        identity = driver._driver_identity()
        self.assertEqual(identity["path"], str((ROOT / "tools/native_qualification/drawingml_compile_driver.py").resolve()))
        self.assertEqual(len(identity["sha256"]), 64)
        self.assertGreater(identity["bytes"], 0)


if __name__ == "__main__":
    unittest.main()
