#!/usr/bin/env python3
"""Source-only tests for the signed consumer preparation driver."""

from __future__ import annotations

import importlib.util
from pathlib import Path
import tempfile
import unittest

HERE = Path(__file__).resolve().parent
SPEC = importlib.util.spec_from_file_location(
    "prepare_signed", HERE / "prepare_signed_native_consumer.py"
)
assert SPEC and SPEC.loader
PREPARE = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(PREPARE)


class SignedConsumerPreparationTests(unittest.TestCase):
    def test_fixture_manifest_covers_all_signed_formats(self) -> None:
        value = PREPARE.fixture_manifest()
        self.assertEqual(
            {row["format"] for row in value["fixtures"]},
            {"DOCX", "PPTX", "XLSX", "XLSB", "DOC", "PPT", "XLS"},
        )

    def test_f2_adapter_keeps_old_shapes_explicit(self) -> None:
        adapter = PREPARE.adapter_source("f2")
        self.assertIn("commit.file", adapter)
        self.assertIn("commit.patch", adapter)
        self.assertIn("commit.diagnostics", adapter)
        self.assertIn("edit_with_limits", adapter)
        self.assertIn("edit_cell_values", adapter)
        self.assertNotIn(".resolve()", adapter)

    def test_final_adapter_uses_triads_and_options(self) -> None:
        adapter = PREPARE.adapter_source("472b")
        self.assertIn("commit.file()", adapter)
        self.assertIn("commit.patch()", adapter)
        self.assertIn("commit.diagnostics()", adapter)
        self.assertIn("edit_with_options", adapter)
        self.assertIn(".resolve()", adapter)
        self.assertNotIn("commit.file;", adapter)

    def test_harness_features_are_bound_to_reviewed_endpoint_projection(self) -> None:
        before = PREPARE._reviewed_harness_features("f2")
        after = PREPARE._reviewed_harness_features("472b")
        self.assertIn("default", before)
        self.assertNotIn("sign", before)
        self.assertIn("sign", after)
        self.assertNotEqual(before, after)
        self.assertIn('features = ["default"', PREPARE.project_toml(Path("/tmp/source"), "f2"))

    def test_unsafe_external_outputs_are_rejected(self) -> None:
        with self.assertRaises(PREPARE.PreparationError):
            PREPARE.safe_external(HERE.parents[3] / "docs/report/attempt", "output")
        with self.assertRaises(PREPARE.PreparationError):
            PREPARE.safe_external(Path("/tmp/3rdparty/output"), "output")

    def test_plan_is_read_only_and_records_caps(self) -> None:
        source = Path(
            "/home/zhuhe/code/litchi-worktrees/.coord/"
            "final-api-capture-472b670547/source"
        )
        if not source.is_dir():
            self.skipTest("final source endpoint is not present")
        with tempfile.TemporaryDirectory(prefix="signed-plan-test-") as directory:
            root = Path(directory)
            before = set(root.iterdir())
            value = PREPARE.plan(source, "472b", root / "out", root / "fixtures")
            self.assertEqual(before, set(root.iterdir()))
            self.assertEqual(value["build_policy"]["jobs"], 1)
            self.assertEqual(value["build_policy"]["memory_cap"], "20G")
            self.assertEqual(value["build_policy"]["target_cap_bytes"], 20_000_000_000)
            self.assertFalse(value["adapter"]["f2_field_bridge"])

    def test_endpoint_capabilities_require_probe_receipts(self) -> None:
        with tempfile.TemporaryDirectory(prefix="signed-capability-test-") as directory:
            root = Path(directory)
            binary = root / "probe"
            binary.write_bytes(b"probe")
            binary.chmod(0o755)
            capabilities = root / "capabilities.json"
            capabilities.write_bytes(
                PREPARE.canonical(
                    {
                        "schema": "litchi_signed_native_capabilities.v1",
                        "source_commit": PREPARE.FINAL_COMMIT,
                        "binary": PREPARE.artifact(binary),
                        "routes": {route: "available" for route in PREPARE.ROUTES},
                        "evidence": {},
                    }
                )
            )
            with self.assertRaises(PREPARE.PreparationError):
                PREPARE._capability_evidence(
                    capabilities,
                    source_commit=PREPARE.FINAL_COMMIT,
                    binary=PREPARE.artifact(binary),
                    fixture_manifest={"path": "/tmp/fixtures.json", "bytes": 1, "sha256": "0" * 64},
                    fixture_root="/tmp/fixtures",
                    endpoint_role="after",
                )


if __name__ == "__main__":
    unittest.main()
