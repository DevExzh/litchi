import json
from pathlib import Path
import tempfile
import unittest

import ppt_lazy_native_qualification as protocol


ROOT = Path(__file__).resolve().parent.parent
MANIFEST = ROOT / "tools" / "native_qualification" / "manifests" / "ppt-lazy-edit.json"


class PptLazyNativeProtocolTests(unittest.TestCase):
    def test_parameterized_manifest_plan_is_build_free(self):
        manifest = protocol._load_manifest(MANIFEST, plan=True)
        self.assertEqual(manifest["baseline"]["ref"], protocol.BASELINE_REF)
        self.assertEqual(manifest["baseline"]["tree"], protocol.BASELINE_TREE)
        self.assertEqual(manifest["baseline"]["lock"]["bytes"], protocol.BASELINE_LOCK_BYTES)
        self.assertEqual(manifest["baseline"]["lock"]["sha256"], protocol.BASELINE_LOCK_SHA256)
        self.assertIsNone(manifest["candidate"])
        self.assertEqual(tuple(protocol.ROUTES), (
            "primary_entry_drop",
            "primary_noop_commit",
            "metadata_entry_drop",
            "metadata_noop_commit",
        ))

    def test_plan_writes_only_a_small_external_receipt(self):
        manifest = protocol._load_manifest(MANIFEST, plan=True)
        with tempfile.TemporaryDirectory(prefix="ppt-lazy-plan-") as temporary:
            output = Path(temporary) / "plan"
            output.mkdir()
            protocol._plan(manifest, MANIFEST, output)
            plan = json.loads((output / "plan.json").read_text(encoding="utf-8"))
            plan_bytes = (output / "plan.json").stat().st_size
        self.assertEqual(plan["status"], "plan_only")
        self.assertEqual(plan["routes"], list(protocol.ROUTES))
        self.assertEqual(plan["abba_order"], list(protocol.ABBA_ORDER))
        self.assertLess(plan_bytes, 1 << 20)

    def test_execution_rejects_unfrozen_candidate(self):
        manifest = protocol._load_manifest(MANIFEST, plan=True)
        with tempfile.TemporaryDirectory(prefix="ppt-lazy-manifest-") as temporary:
            path = Path(temporary) / "manifest.json"
            path.write_text(json.dumps(manifest), encoding="utf-8")
            with self.assertRaises(protocol.RunFailure):
                protocol._load_manifest(path, plan=False)

    def test_source_only_protocol_has_fixed_caps(self):
        manifest = protocol._load_manifest(MANIFEST, plan=True)
        protocol_value = manifest["protocol"]
        self.assertEqual(protocol_value["blocks"], 2)
        self.assertEqual(protocol_value["warmups"], 3)
        self.assertGreaterEqual(protocol_value["target_ms"], 500)
        self.assertGreaterEqual(protocol_value["minimum_ms"], 150)
        self.assertEqual(protocol.TARGET_HARD_LIMIT_BYTES, 20_000_000_000)
        self.assertEqual(protocol.MEMORY_HARD_LIMIT_BYTES, 20_000_000_000)
        self.assertEqual(protocol.MAX_CHILD_OUTPUT, 4 * 1024 * 1024)

    def test_timed_validator_enforces_minimum_sample(self):
        manifest = protocol._load_manifest(MANIFEST, plan=True)
        allocation = {"status": "unavailable", "scope": "operation_global_system_allocator"}
        report_base = {
            "schema": protocol.PROBE_SCHEMA,
            "mode": "timed",
            "route": "primary_entry_drop",
        }
        short = dict(report_base, samples=[{"elapsed_ns": 149_999_999, "iterations": 1, "allocation": allocation}])
        with self.assertRaises(protocol.RunFailure):
            protocol._validate_timed(
                short,
                "primary_entry_drop",
                1,
                False,
                manifest["protocol"],
                "short",
            )
        valid = dict(report_base, samples=[{"elapsed_ns": 150_000_000, "iterations": 1, "allocation": allocation}])
        result = protocol._validate_timed(
            valid,
            "primary_entry_drop",
            1,
            False,
            manifest["protocol"],
            "valid",
        )
        self.assertEqual(result["samples"][0]["elapsed_ns"], 150_000_000)

    def test_allocator_validator_requires_live_byte_conservation(self):
        manifest = protocol._load_manifest(MANIFEST, plan=True)
        allocation = {
            "status": "measured",
            "scope": "operation_global_system_allocator",
            "allocation_calls": 1,
            "deallocation_calls": 0,
            "reallocation_calls": 0,
            "failed_allocation_calls": 0,
            "allocated_bytes": 8,
            "deallocated_bytes": 0,
            "live_bytes_before": 10,
            "live_bytes_after": 10,
            "peak_live_bytes_before": 10,
            "peak_live_bytes_after": 18,
        }
        report = {
            "schema": protocol.PROBE_SCHEMA,
            "mode": "timed",
            "route": "primary_entry_drop",
            "samples": [{"elapsed_ns": 150_000_000, "iterations": 1, "allocation": allocation}],
        }
        with self.assertRaises(protocol.RunFailure):
            protocol._validate_timed(
                report,
                "primary_entry_drop",
                1,
                True,
                manifest["protocol"],
                "broken-allocation",
            )


if __name__ == "__main__":
    unittest.main()
