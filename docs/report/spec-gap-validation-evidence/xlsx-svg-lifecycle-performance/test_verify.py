"""Adversarial receipt checks; these are not performance measurements."""

import json
import hashlib
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import patch

import verify


class ReceiptChecks(unittest.TestCase):
    def setUp(self):
        self.directory = TemporaryDirectory(prefix="litchi-profile-verifier-")
        self.addCleanup(self.directory.cleanup)
        self.root = Path(self.directory.name)

    def process(self, name="capture-p1.json"):
        path = self.root / name
        path.with_suffix(".stderr.log").write_bytes(b"")
        path.with_suffix(".time.txt").write_text("\tExit status: 0\n")
        return path

    def test_process_requires_empty_stderr_and_successful_exit(self):
        path = self.process()
        verify.verify_process_output(path)
        for stderr, timing in [
            (b"debug output\n", "Exit status: 0\n"),
            (b"", "Exit status: 1\n"),
            (b"", "Command terminated by signal 9\n"),
            (b"", "Exit status: 0\nExit status: 0\n"),
        ]:
            with self.subTest(stderr=stderr, timing=timing):
                path.with_suffix(".stderr.log").write_bytes(stderr)
                path.with_suffix(".time.txt").write_text(timing)
                with self.assertRaises(AssertionError):
                    verify.verify_process_output(path)

    def test_missing_stderr_is_not_evidence_of_silence(self):
        path = self.process()
        path.with_suffix(".stderr.log").unlink()
        with self.assertRaisesRegex(AssertionError, "stderr receipt missing"):
            verify.verify_process_output(path)

    def test_binary_identity_must_be_well_formed_and_stable(self):
        receipt = "a" * 64 + "  /removed/disposable/binary\n"
        (self.root / "binary.sha256").write_text(receipt)
        (self.root / "binary-after.sha256").write_text(receipt)
        self.assertEqual(verify.verify_binary_receipts(self.root), "a" * 64)
        (self.root / "binary-after.sha256").write_text(receipt.replace("aaa", "bbb", 1))
        with self.assertRaisesRegex(AssertionError, "changed during measurement"):
            verify.verify_binary_receipts(self.root)
        for name in ("binary.sha256", "binary-after.sha256"):
            (self.root / name).write_text("invalid  /binary\n")
        with self.assertRaisesRegex(AssertionError, "malformed executable SHA"):
            verify.verify_binary_receipts(self.root)

    def test_three_arbitrary_filenames_do_not_prove_process_identities(self):
        for number in (1, 2, 4):
            (self.root / f"capture-p{number}.json").write_text("{}")
        with patch.object(verify, "LANES", ("capture",)):
            with self.assertRaisesRegex(AssertionError, "process identities"):
                verify.verify_lanes(self.root)

    def test_input_size_must_match_even_when_digests_match(self):
        sample = {
            "expected_success": True, "actual_success": True,
            "semantic_ok": True, "output_exact": True,
            "direct_allocated_bytes": 0, "realloc_new_bytes": 0,
            "realloc_old_bytes": 0, "deallocated_bytes": 0,
            "live_before": 0, "live_after": 0, "requested_alloc_bytes": 0,
            "alloc_balance_ok": True, "alloc_invalid": False,
            "alloc_failed": 0, "error": None,
        }
        receipt = {
            "schema": "xlsx-svg-lifecycle-profile-v1", "lane": "capture",
            "warmup": 2, "sample_count": 20, "samples": [sample] * 20,
            "expected_success": True, "input_hash_fnv1a64": 1,
            "input_sha256": "a" * 64, "input_bytes": 100,
        }
        for number in range(1, 4):
            self.process(f"capture-p{number}.json").write_text(json.dumps(receipt))
        with patch.object(verify, "LANES", ("capture",)):
            verify.verify_lanes(self.root)
            receipt["input_bytes"] = 101
            (self.root / "capture-p3.json").write_text(json.dumps(receipt))
            with self.assertRaisesRegex(AssertionError, "fixture size changed"):
                verify.verify_lanes(self.root)

    def test_native_receipt_identity_must_match_documented_producer_fixture(self):
        native = self.root / "native.xlsx"
        native.write_bytes(b"bounded synthetic verifier input")
        digest = hashlib.sha256(native.read_bytes()).hexdigest()
        corpus = {"source_fixture": {"path": native.name, "sha256": digest}}
        receipt = {"input_sha256": digest, "input_bytes": native.stat().st_size}
        for number in range(1, 4):
            (self.root / f"capture_native_fixture-p{number}.json").write_text(json.dumps(receipt))
        verify.verify_native_identity(self.root, corpus, self.root)
        receipt["input_sha256"] = "b" * 64
        (self.root / "capture_native_fixture-p2.json").write_text(json.dumps(receipt))
        with self.assertRaisesRegex(AssertionError, "different fixture"):
            verify.verify_native_identity(self.root, corpus, self.root)

    def test_namespace_boundary_requires_explicit_first_refused_active_count(self):
        path = self.root / "namespace_limit_refusal-p1.json"
        receipt = {
            "namespace_generated_bindings": 16378,
            "namespace_active_bindings": 16385,
            "namespace_active_limit": 16384,
        }
        self.assertEqual(
            verify.verify_namespace_boundary(receipt, "namespace_limit_refusal", path),
            (16378, 16385, 16384),
        )
        for field in receipt:
            for bad in (None, True, "16384", 0, -1):
                with self.subTest(field=field, bad=bad):
                    with self.assertRaisesRegex(AssertionError, "missing or malformed"):
                        verify.verify_namespace_boundary(
                            {**receipt, field: bad}, "namespace_limit_refusal", path
                        )
        for replacement in (
            {"namespace_active_bindings": 16378},
            {"namespace_active_bindings": 16386},
        ):
            with self.assertRaises(AssertionError):
                verify.verify_namespace_boundary({**receipt, **replacement}, "namespace_limit_refusal", path)
        with self.assertRaisesRegex(AssertionError, "unexpected namespace boundary"):
            verify.verify_namespace_boundary(receipt, "capture", path)

    def test_namespace_boundary_must_match_across_processes(self):
        lane = "namespace_limit_refusal"
        sample = {
            "expected_success": False, "actual_success": False,
            "semantic_ok": True, "output_exact": True,
            "direct_allocated_bytes": 0, "realloc_new_bytes": 0,
            "realloc_old_bytes": 0, "deallocated_bytes": 0,
            "live_before": 0, "live_after": 0, "requested_alloc_bytes": 0,
            "alloc_balance_ok": True, "alloc_invalid": False,
            "alloc_failed": 0, "error": {"class": "namespace_limit"},
        }
        receipt = {
            "schema": "xlsx-svg-lifecycle-profile-v1", "lane": lane,
            "warmup": 2, "sample_count": 20, "samples": [sample] * 20,
            "expected_success": False, "input_hash_fnv1a64": 1,
            "input_sha256": "a" * 64, "input_bytes": 100,
            "namespace_generated_bindings": 16378,
            "namespace_active_bindings": 16385,
            "namespace_active_limit": 16384,
        }
        for number in range(1, 4):
            self.process(f"{lane}-p{number}.json").write_text(json.dumps(receipt))
        with patch.object(verify, "LANES", (lane,)):
            verify.verify_lanes(self.root)
            receipt["namespace_generated_bindings"] = 16379
            receipt["namespace_active_bindings"] = 16386
            receipt["namespace_active_limit"] = 16385
            self.process(f"{lane}-p3.json").write_text(json.dumps(receipt))
            with self.assertRaisesRegex(AssertionError, "namespace boundary changed"):
                verify.verify_lanes(self.root)

    def test_svg_caller_ceiling_requires_exact_integer(self):
        path = self.root / "limit_small-p1.json"
        profile = {"kind": "svg_input_bytes", "max": 32 * 1024 * 1024}
        self.assertEqual(
            verify.verify_caller_limits(
                {"caller_limits": profile}, "limit_small", path
            ),
            ("svg_input_bytes", 32 * 1024 * 1024),
        )
        for bad in (None, 32 * 1024 * 1024.0, True, 0, -1, "33554432"):
            with self.subTest(bad=bad):
                with self.assertRaisesRegex(AssertionError, "input ceiling mismatch"):
                    verify.verify_caller_limits(
                        {"caller_limits": {"kind": "svg_input_bytes", "max": bad}},
                        "limit_small",
                        path,
                    )

    def test_composite_caller_ceilings_reject_missing_float_bool_and_nonpositive_values(self):
        path = self.root / "mixed_caps_rejection-p1.json"
        valid = [
            {"name": name, "value": index + 1}
            for index, name in enumerate(verify.COMPOSITE_LIMIT_NAMES)
        ]
        profile = {"kind": "composite_read_limits", "ceilings": valid}
        self.assertEqual(
            verify.verify_caller_limits({"caller_limits": profile}, "mixed_caps_rejection", path),
            (
                "composite_read_limits",
                tuple((name, index + 1) for index, name in enumerate(verify.COMPOSITE_LIMIT_NAMES)),
            ),
        )
        for bad in (None, 1.0, True, 0, -1):
            with self.subTest(bad=bad):
                ceilings = [dict(item) for item in valid]
                ceilings[0]["value"] = bad
                with self.assertRaisesRegex(AssertionError, "ceiling value malformed"):
                    verify.verify_caller_limits(
                        {"caller_limits": {"kind": "composite_read_limits", "ceilings": ceilings}},
                        "mixed_caps_rejection",
                        path,
                    )
        for replacement in (None, [], valid[:-1], valid + [{"name": "extra", "value": 1}]):
            with self.subTest(replacement=replacement):
                profile_value = {"kind": "composite_read_limits", "ceilings": replacement}
                message = "composite caller ceilings missing" if replacement is None else "composite caller ceiling count changed"
                with self.assertRaisesRegex(AssertionError, message):
                    verify.verify_caller_limits(
                        {"caller_limits": profile_value}, "mixed_caps_rejection", path
                    )

    def test_composite_caller_ceiling_must_match_across_processes(self):
        lane = "mixed_caps_rejection"
        sample = {
            "expected_success": False,
            "actual_success": False,
            "semantic_ok": True,
            "output_exact": True,
            "direct_allocated_bytes": 0,
            "realloc_new_bytes": 0,
            "realloc_old_bytes": 0,
            "deallocated_bytes": 0,
            "live_before": 0,
            "live_after": 0,
            "requested_alloc_bytes": 0,
            "alloc_balance_ok": True,
            "alloc_invalid": False,
            "alloc_failed": 0,
            "error": {"class": "mixed_limit"},
        }
        profile = {
            "kind": "composite_read_limits",
            "ceilings": [
                {"name": name, "value": index + 1}
                for index, name in enumerate(verify.COMPOSITE_LIMIT_NAMES)
            ],
        }
        receipt = {
            "schema": "xlsx-svg-lifecycle-profile-v1",
            "lane": lane,
            "warmup": 2,
            "sample_count": 20,
            "samples": [sample] * 20,
            "expected_success": False,
            "input_hash_fnv1a64": 1,
            "input_sha256": "a" * 64,
            "input_bytes": 100,
            "caller_limits": profile,
        }
        for number in range(1, 4):
            self.process(f"{lane}-p{number}.json").write_text(json.dumps(receipt))
        with patch.object(verify, "LANES", (lane,)):
            verify.verify_lanes(self.root)
            changed = json.loads(json.dumps(receipt))
            changed["caller_limits"]["ceilings"][0]["value"] = 2
            self.process(f"{lane}-p3.json").write_text(json.dumps(changed))
            with self.assertRaisesRegex(AssertionError, "caller-limit profile changed"):
                verify.verify_lanes(self.root)



if __name__ == "__main__":
    unittest.main()
