"""Source-only contract tests for the supplemental cold PPT route."""
from __future__ import annotations
import hashlib
import importlib.util
import json
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest
ROOT = Path(__file__).resolve().parents[1]
SPEC = importlib.util.spec_from_file_location(
    "ppt_lazy_cold_first_commit", ROOT / "tools/ppt_lazy_cold_first_commit.py"
)
assert SPEC and SPEC.loader
MODULE = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(MODULE)
class ColdFirstCommitTests(unittest.TestCase):
    def test_route_is_separate_and_cold(self) -> None:
        self.assertEqual(MODULE.ROUTE, "primary_first_commit_cold")
        self.assertEqual(MODULE.UNDERLYING_ROUTE, "primary_noop_commit")
        self.assertEqual(MODULE.CPU, 29)
        self.assertIn("--warmups", MODULE.ARGS)
        self.assertEqual(MODULE.ARGS[MODULE.ARGS.index("--warmups") + 1], "0")
        self.assertEqual(MODULE.ARGS[MODULE.ARGS.index("--iterations") + 1], "1")
    def test_source_only_plan_binds_reviewed_binary_and_fixture(self) -> None:
        with tempfile.TemporaryDirectory(prefix="ppt-cold-plan-") as temporary:
            root = Path(temporary)
            binary = root / "ppt_lazy_native"
            fixture = root / "fixture.ppt"
            binary.write_bytes(b"reviewed binary")
            fixture.write_bytes(b"fixture")
            # Use the fixed identity table only to test that callers cannot
            # supply an arbitrary expected hash; no execution is performed.
            role = "candidate-normal"
            expected_binary = MODULE.BINARY_IDENTITIES[role]
            binary.write_bytes(b"x" * expected_binary["bytes"])
            os.chmod(binary, 0o755)
            digest = hashlib.sha256(binary.read_bytes()).hexdigest()
            MODULE.BINARY_IDENTITIES[role]["sha256"] = digest
            MODULE.FIXTURE["bytes"] = fixture.stat().st_size
            MODULE.FIXTURE["sha256"] = hashlib.sha256(fixture.read_bytes()).hexdigest()
            try:
                result = MODULE.run(binary, fixture, role, execute=False)
                self.assertEqual(result["route"], "primary_first_commit_cold")
                self.assertEqual(result["command"][result["command"].index("--warmups") + 1], "0")
                self.assertNotIn("report", result)
            finally:
                MODULE.BINARY_IDENTITIES[role]["sha256"] = "ac72404c98bf37f5a7b9433c01880547a4b34d24a8d346dc359ecc765770ff1a"
                MODULE.FIXTURE["bytes"] = 14_848
                MODULE.FIXTURE["sha256"] = "02fa44ca76903943e80a3a2da1f9dbd326c7a4a047540c3aa9cf211e34ef617c"
    def test_report_requires_one_cold_iteration_and_exact_noop(self) -> None:
        fixture = dict(MODULE.FIXTURE)
        proof = {
            "output": fixture,
            "source_exact": True,
            "changed": False,
            "patch_noop": True,
            "operation_count": 0,
            "patch_before_exact": True,
            "patch_after_exact": True,
            "entry": "commit",
            "fixture": {"key": "default", "output": fixture},
        }
        report = {
            "schema": MODULE.PROBE_SCHEMA,
            "mode": "timed",
            "route": MODULE.UNDERLYING_ROUTE,
            "samples": [{
                "elapsed_ns": 1,
                "iterations": 1,
                "allocation": {"status": "unavailable", "scope": "operation_global_system_allocator"},
            }],
            "fixture": {"key": "default", "output": fixture},
            "postcondition": proof,
        }
        self.assertEqual(MODULE.validate_report(report, fixture, allocator=False)["route"], MODULE.UNDERLYING_ROUTE)
        report["samples"][0]["iterations"] = 2
        with self.assertRaises(MODULE.ColdRouteError):
            MODULE.validate_report(report, fixture, allocator=False)
    def test_strict_json_rejects_duplicate_and_nonfinite_values(self) -> None:
        with self.assertRaises(MODULE.ColdRouteError):
            MODULE._strict_json(b'{"x": 1, "x": 2}', "test")
        with self.assertRaises(MODULE.ColdRouteError):
            MODULE._strict_json(b'{"x": 1e999}', "test")

    def test_public_cli_exposes_symmetric_pair_contract(self) -> None:
        completed = subprocess.run(
            [sys.executable, str(ROOT / "tools" / "ppt_lazy_cold_first_commit.py"), "--help"],
            check=False,
            capture_output=True,
            text=True,
        )
        self.assertEqual(completed.returncode, 0)
        self.assertIn("--manifest", completed.stdout)
        self.assertIn("--projection", completed.stdout)
        self.assertIn("--out-dir", completed.stdout)
        self.assertIn("--rounds", completed.stdout)

    def test_pair_report_aggregates_operation_elapsed_ns_in_abba_order(self) -> None:
        rows = []
        for round_index in range(1):
            for role in MODULE.ABBA_ORDER:
                for kind in ("normal", "allocator"):
                    rows.append({
                        "round": round_index,
                        "role": role,
                        "endpoint": MODULE.ROLE_ENDPOINT[role],
                        "kind": kind,
                        "sample": {
                            "elapsed_ns": 10 if kind == "normal" else 20,
                            "iterations": 1,
                            "allocation": {"status": "unavailable"},
                        },
                    })
        report = {
            "schema": MODULE.SCHEMA,
            "route": MODULE.ROUTE,
            "abba_order": list(MODULE.ABBA_ORDER),
            "rounds": 1,
            "samples": rows,
            "aggregate": {
                "baseline": {
                    "normal": {"operation_total_ns": 20, "operation_count": 2, "operation_mean_ns": 10},
                    "allocator": {"operation_total_ns": 40, "operation_count": 2, "operation_mean_ns": 20},
                },
                "candidate": {
                    "normal": {"operation_total_ns": 20, "operation_count": 2, "operation_mean_ns": 10},
                    "allocator": {"operation_total_ns": 40, "operation_count": 2, "operation_mean_ns": 20},
                },
            },
        }
        self.assertNotIn("median", MODULE.validate_pair_report(report, rounds=1))
        full = dict(
            report,
            underlying_route=MODULE.UNDERLYING_ROUTE,
            warmups=0,
            iterations=1,
            commands={},
            manifest={"path": "manifest"},
            build_input_projection={"path": "projection"},
            claim_scope="operation only",
            execution="native_executed",
        )
        self.assertEqual(MODULE.validate_pair_report(full, rounds=1)["execution"], "native_executed")
        reordered = dict(report, samples=list(report["samples"]))
        reordered["samples"][0], reordered["samples"][2] = reordered["samples"][2], reordered["samples"][0]
        with self.assertRaises(MODULE.ColdRouteError):
            MODULE.validate_pair_report(reordered, rounds=1)
        report["aggregate"]["candidate"]["normal"]["operation_total_ns"] = 1
        with self.assertRaises(MODULE.ColdRouteError):
            MODULE.validate_pair_report(report, rounds=1)
if __name__ == "__main__":
    unittest.main()
