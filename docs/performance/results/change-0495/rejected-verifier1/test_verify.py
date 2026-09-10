#!/usr/bin/env python3
"""Fixture-only tests for the 0495 bundle verifier."""

from __future__ import annotations

import os
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

import measure
import verify as bundle


def _manifest() -> dict[str, object]:
    return {
        "schema": bundle.FINAL_GATES_SCHEMA,
        "version": bundle.FINAL_GATES_VERSION,
        "commands": {
            f"final-gate-{index + 1}": list(command)
            for index, command in enumerate(bundle._expected_final_commands())
        },
    }


def _allocation() -> dict[str, object]:
    return {
        "status": "measured", "scope": measure.SAMPLE_ALLOCATION_SCOPE,
        "allocation_calls": 4, "deallocation_calls": 0,
        "reallocation_calls": 1, "failed_allocation_calls": 0,
        "allocated_bytes": 100, "deallocated_bytes": 0,
        "live_bytes_before": 10, "live_bytes_after": 110,
        "peak_live_bytes_before": 20, "peak_live_bytes_after": 120,
        "region_peak_live_bytes": 115,
    }


class VerifyFixtures(unittest.TestCase):
    def test_exact_command_set_accepts_harness_and_docx_format_files(self) -> None:
        pairs = bundle._validate_final_gates(_manifest())
        self.assertEqual(len(pairs), 17)
        format_commands = [command for _, command in pairs if command and command[0] == "rustfmt"]
        self.assertEqual(len(format_commands), 2)
        self.assertEqual(tuple(format_commands[0][7:]), bundle.FORMAT_FILES)
        self.assertEqual(len(bundle.DOCX_FORMAT_FILES), 26)
        self.assertIn("crates/litchi-docx/tests/source_backed_semantic.rs",
                      bundle.DOCX_FORMAT_FILES)
        self.assertEqual(len(bundle.DOCX_FORMAT_FILES + bundle.OPC_FORMAT_FILES), 27)
        self.assertEqual(tuple(format_commands[1][7:]),
                         bundle.DOCX_FORMAT_FILES + bundle.OPC_FORMAT_FILES)

    def test_exact_command_set_rejects_missing_or_arbitrary_gate(self) -> None:
        missing = _manifest()
        missing["commands"].pop("final-gate-1")  # type: ignore[union-attr]
        with self.assertRaises(measure.ProviderMatrixError):
            bundle._validate_final_gates(missing)
        arbitrary = _manifest()
        arbitrary["commands"]["final-gate-1"] = ["true"]  # type: ignore[index]
        with self.assertRaises(measure.ProviderMatrixError):
            bundle._validate_final_gates(arbitrary)

    def test_workspace_fixture_custody_includes_docx_compile_inputs(self) -> None:
        custody = bundle._validate_baseline_inputs()
        self.assertEqual(len(custody["fixtures"]), 2)
        self.assertEqual(len(custody["harness_runtime_fixtures"]), 2)
        self.assertEqual(set(custody["harness_runtime_fixtures"]),
                         bundle.HARNESS_RUNTIME_FIXTURE_NAMES)
        self.assertEqual(len(custody["docx_compile_fixtures"]), 15)
        self.assertEqual(len(custody["docx_runtime_fixtures"]), 63)
        self.assertEqual(len(custody["opc_compile_fixtures"]), 7)
        self.assertEqual(set(custody["opc_compile_fixtures"]),
                         bundle.OPC_COMPILE_FIXTURE_NAMES)
        self.assertIn(
            "docs/performance/results/change-0416/corpus/opc-local-only-signed.zip",
            custody["opc_compile_fixtures"],
        )
        self.assertEqual(len(custody["opc_runtime_fixtures"]), 2)
        self.assertEqual(set(custody["opc_runtime_fixtures"]),
                         bundle.OPC_RUNTIME_FIXTURE_NAMES)
        self.assertEqual(len(custody["boundary_check_inputs"]), 2)
        self.assertEqual(len(custody["boundary_adr_inputs"]), 30)

    def test_opc_fixture_custody_rejects_omitting_external_signed_input(self) -> None:
        original_json = bundle._json

        def omit_signed_input(path: Path, label: str) -> object:
            value = original_json(path, label)
            if Path(path).name == "opc-compile-fixture-inputs.json":
                value = dict(value)
                value.pop(
                    "docs/performance/results/change-0416/corpus/opc-local-only-signed.zip",
                    None,
                )
            return value

        with patch.object(bundle, "_json", side_effect=omit_signed_input):
            with self.assertRaises(measure.ProviderMatrixError):
                bundle._validate_baseline_inputs()

    def test_opc_runtime_fixture_custody_rejects_omitting_runtime_input(self) -> None:
        original_json = bundle._json

        def omit_runtime_input(path: Path, label: str) -> object:
            value = original_json(path, label)
            if Path(path).name == "opc-runtime-fixture-inputs.json":
                value = dict(value)
                value.pop("test-data/poi/test-data/openxml4j/50154.xlsx", None)
            return value

        with patch.object(bundle, "_json", side_effect=omit_runtime_input):
            with self.assertRaises(measure.ProviderMatrixError):
                bundle._validate_baseline_inputs()

    def test_harness_runtime_fixture_custody_rejects_omitting_runtime_input(self) -> None:
        original_json = bundle._json

        def omit_runtime_input(path: Path, label: str) -> object:
            value = original_json(path, label)
            if Path(path).name == "harness-runtime-fixture-inputs.json":
                value = dict(value)
                value.pop("test-data/ooxml/pptx/shapes.pptx", None)
            return value

        with patch.object(bundle, "_json", side_effect=omit_runtime_input):
            with self.assertRaises(measure.ProviderMatrixError):
                bundle._validate_baseline_inputs()

    def test_gate_source_binding_can_be_skipped_only_for_helper_gate(self) -> None:
        command = list(bundle._expected_final_commands()[5])
        with tempfile.TemporaryDirectory(prefix="litchi-bundle-0495-") as raw:
            path = Path(raw) / "helper-gate.json"
            path.write_text('{"label":"helper-gate"}\n', encoding="utf-8")
            binding = {"source": {"current": "postprocessed"}, "argv": command}
            with patch.object(bundle.measure, "_gate_binding", return_value=binding):
                self.assertEqual(
                    bundle._validate_gate_receipt("helper-gate", command, path), binding
                )

    def test_rust_gate_requires_frozen_source_binding(self) -> None:
        command = list(bundle._expected_final_commands()[0])
        with tempfile.TemporaryDirectory(prefix="litchi-bundle-0495-") as raw:
            path = Path(raw) / "rust-gate.json"
            path.write_text('{"label":"rust-gate"}\n', encoding="utf-8")
            binding = {"source": {"current": "drifted"}, "argv": command}
            with patch.object(bundle.measure, "_gate_binding", return_value=binding):
                with self.assertRaises(measure.ProviderMatrixError):
                    bundle._validate_gate_receipt("rust-gate", command, path, {"frozen": "source"})

    def test_allocator_conservation_rejects_live_balance_tamper(self) -> None:
        allocation = _allocation()
        bundle._validate_allocation(allocation, "fixture.allocation")
        allocation["live_bytes_after"] = 111
        with self.assertRaises(measure.ProviderMatrixError):
            bundle._validate_allocation(allocation, "fixture.allocation")

    def test_allocator_requires_strict_uint_and_peak_bounds(self) -> None:
        for field, value in (("allocated_bytes", True), ("peak_live_bytes_before", 9),
                             ("region_peak_live_bytes", 121), ("allocation_calls", 0)):
            allocation = _allocation()
            allocation[field] = value
            with self.subTest(field=field), self.assertRaises(measure.ProviderMatrixError):
                bundle._validate_allocation(allocation, "fixture.allocation")

    def test_reallocation_does_not_require_deallocation_calls(self) -> None:
        allocation = _allocation()
        allocation.update({"deallocation_calls": 0, "allocated_bytes": 100,
                           "deallocated_bytes": 0, "live_bytes_after": 110,
                           "peak_live_bytes_after": 120, "region_peak_live_bytes": 115})
        bundle._validate_allocation(allocation, "fixture.allocation")

    def test_inventory_rejects_special_files_and_excludes_seal(self) -> None:
        with tempfile.TemporaryDirectory(prefix="litchi-bundle-0495-") as raw:
            root = Path(raw)
            (root / "evidence.json").write_text("{}\n", encoding="utf-8")
            (root / "seal.json").write_text("{}\n", encoding="utf-8")
            self.assertEqual(bundle.inventory(root), {"evidence.json": bundle.meta(root / "evidence.json")})
            os.mkfifo(root / "unsafe.fifo")
            with self.assertRaises(RuntimeError):
                bundle.inventory(root)

    def test_historical_receipt_requires_archived_support_helper(self) -> None:
        # The production history validator rejects a made-up helper hash before
        # it can be mistaken for an authenticated development receipt.
        with tempfile.TemporaryDirectory(prefix="litchi-bundle-0495-") as raw:
            path = Path(raw) / "development.json"
            path.write_text("{}\n", encoding="utf-8")
            with self.assertRaises(measure.ProviderMatrixError):
                bundle._validate_historical_receipt(path)


if __name__ == "__main__":
    raise SystemExit(unittest.main())
