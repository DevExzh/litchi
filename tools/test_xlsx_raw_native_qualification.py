#!/usr/bin/env python3
"""Static tests for the guarded XLSX raw native qualification protocol."""

from __future__ import annotations

import hashlib
import json
from pathlib import Path
import sys
import tempfile
import unittest


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))
from native_qualification import xlsx_raw_identity_validator as validator  # noqa: E402


def _allocation_sample(status: str) -> dict[str, int | str]:
    if status == "unavailable":
        return {"status": status, "scope": "operation_global_system_allocator"}
    return {
        "status": status,
        "scope": "operation_global_system_allocator",
        "allocation_calls": 1,
        "deallocation_calls": 1,
        "reallocation_calls": 0,
        "failed_allocation_calls": 0,
        "allocated_bytes": 1,
        "deallocated_bytes": 1,
        "live_bytes_before": 0,
        "live_bytes_after": 0,
        "peak_live_bytes_before": 1,
        "peak_live_bytes_after": 1,
        "region_peak_live_bytes": 1,
    }


def _report(binary: Path, route: str, variant: str, revision: str) -> dict[str, object]:
    binary_bytes = binary.read_bytes()
    sample_status = "measured" if route == "alloc" else "unavailable"
    phase = {
        "elapsed_ns": [0, 1, 2, 3, 4],
        "allocation_samples": [_allocation_sample(sample_status) for _ in range(5)],
    }
    operations = []
    for operation in validator.OPERATION_NAMES:
        operations.append(
            {
                "operation": operation,
                "base_bytes": 1,
                "base_sha256": "a" * 64,
                "output_bytes": 1,
                "output_sha256": "b" * 64,
                "elapsed_ns": [0, 1, 2, 3, 4],
                "allocation_samples": [_allocation_sample(sample_status) for _ in range(5)],
                "native_identity": ["rIdRawProbe0:/xl/rawProbe/part-000.bin"],
                **{field: True for field in validator.FIDELITY_FIELDS},
            }
        )
    return {
        "schema": validator.SCHEMA,
        "api_variant": variant,
        "source_revision": revision,
        "corpus_generator": "litchi-xlsx-raw-identity-package-v2",
        "lifecycle_scope": "owned open, first selected read, raw insertion, native identity",
        "iterations": 5,
        "warmups": 2,
        "binary": {
            "path": str(binary),
            "bytes": len(binary_bytes),
            "sha256": hashlib.sha256(binary_bytes).hexdigest(),
        },
        "cases": [
            {
                "name": name,
                "extra_parts": 0,
                "extra_part_bytes": 0,
                "cache_count": 1,
                "open": phase,
                "first_read": phase,
                "operations": operations,
            }
            for name in validator.CASE_NAMES
        ],
    }


class XlsxRawNativeQualificationTests(unittest.TestCase):
    def _write_valid_run(self, root: Path) -> tuple[Path, Path, Path]:
        output = root / "output"
        before_target = root / "targets" / "before"
        after_target = root / "targets" / "after"
        output.mkdir()
        for target in (before_target, after_target):
            (target / "release").mkdir(parents=True)
            for route in validator.ROUTES:
                binary = target / "release" / f"xlsx-raw-identity-{route}"
                binary.write_bytes(f"{target.name}-{route}".encode())
        for index, leg in enumerate(validator.ABBA_ORDER, start=1):
            variant = "original" if leg == "before" else "candidate"
            revision = validator.BASELINE_REVISION if leg == "before" else "1" * 40
            target = before_target if leg == "before" else after_target
            for route in validator.ROUTES:
                binary = target / "release" / f"xlsx-raw-identity-{route}"
                path = output / f"xlsx-raw-{route}-{leg}-{index}.json"
                path.write_text(json.dumps(_report(binary, route, variant, revision)), encoding="utf-8")
        return output, before_target, after_target

    def test_validates_both_routes_and_all_four_abba_legs(self) -> None:
        with tempfile.TemporaryDirectory(prefix="xlsx-raw-native-qual-") as directory:
            output, before_target, after_target = self._write_valid_run(Path(directory))
            summary = validator.validate_directory(
                output,
                validator.BASELINE_REVISION,
                "1" * 40,
                before_target,
                after_target,
            )
            self.assertEqual(len(summary["reports"]), 8)
            self.assertEqual(summary["routes"], ["alloc", "native"])
            self.assertFalse(summary["performance_claim"])

    def test_rejects_fidelity_loss_in_one_report(self) -> None:
        with tempfile.TemporaryDirectory(prefix="xlsx-raw-native-qual-") as directory:
            output, before_target, after_target = self._write_valid_run(Path(directory))
            report_path = output / "xlsx-raw-native-after-2.json"
            report = json.loads(report_path.read_text(encoding="utf-8"))
            report["cases"][0]["operations"][0]["inverse_restores_source_bytes"] = False
            report_path.write_text(json.dumps(report), encoding="utf-8")
            with self.assertRaisesRegex(validator.ValidationError, "fidelity guard"):
                validator.validate_directory(
                    output,
                    validator.BASELINE_REVISION,
                    "1" * 40,
                    before_target,
                    after_target,
                )

    def test_rejects_non_original_baseline_revision(self) -> None:
        with tempfile.TemporaryDirectory(prefix="xlsx-raw-native-qual-") as directory:
            output, before_target, after_target = self._write_valid_run(Path(directory))
            with self.assertRaisesRegex(validator.ValidationError, "original 0bef"):
                validator.validate_directory(
                    output,
                    "2" * 40,
                    "1" * 40,
                    before_target,
                    after_target,
                )


if __name__ == "__main__":
    unittest.main()
