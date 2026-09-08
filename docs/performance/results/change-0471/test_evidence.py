#!/usr/bin/env python3
"""Focused tests for the 0471 evidence boundary and portable seal."""

from __future__ import annotations

import hashlib
import json
import shutil
import tempfile
import unittest
from pathlib import Path
from unittest import mock

import analyze
import verify


class EvidenceTests(unittest.TestCase):
    def test_frozen_matrix_has_seven_rows(self) -> None:
        protocol = verify.verify_protocol(verify.ROOT)
        self.assertEqual(len(analyze.expected_rows(protocol)), 7)
        self.assertIn(("ppt_fresh_write_to", "payload-heavy"), analyze.expected_rows(protocol))

    def test_projection_only_removes_empty_source_vectors(self) -> None:
        report = {
            "results": [{
                "source": {"read_calls": [], "read_bytes": [7]},
                "operation_metrics": {"values": []},
                "elapsed_ns": {"samples": [1, 2, 3]},
            }],
        }
        projected, removed = analyze.comparison_projection(report)
        self.assertEqual(removed, ["results/0/source/read_calls"])
        self.assertEqual(report["results"][0]["source"]["read_calls"], [])
        self.assertNotIn("read_calls", projected["results"][0]["source"])
        self.assertEqual(projected["results"][0]["source"]["read_bytes"], [7])
        self.assertEqual(projected["results"][0]["operation_metrics"], {"values": []})
        self.assertEqual(projected["results"][0]["elapsed_ns"], {"samples": [1, 2, 3]})

    def test_heaptrack_parser_retains_display_and_excludes_latency(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "heaptrack-print.stdout"
            path.write_text(
                "calls to allocation functions: 10 (20/s)\n"
                "temporary memory allocations: 4 (8/s)\n"
                "peak heap memory consumption: 12.50M\n",
                encoding="utf-8",
            )
            parsed = analyze.parse_heaptrack(path, root=Path(directory))
        self.assertEqual(parsed["allocation_calls"], 10)
        self.assertEqual(parsed["temporary_allocations"], 4)
        self.assertEqual(parsed["peak_heap_display"], "12.50M")
        self.assertEqual(parsed["latency_comparison"], "excluded")

    def test_heaptrack_parser_rejects_missing_peak(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "heaptrack-print.stdout"
            path.write_text(
                "calls to allocation functions: 10\n"
                "temporary memory allocations: 4\n",
                encoding="utf-8",
            )
            with self.assertRaises(analyze.AnalysisError):
                analyze.parse_heaptrack(path, root=Path(directory))

    def test_seal_requires_exact_regular_file_coverage(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            (root / "nested").mkdir()
            (root / "nested" / "a.txt").write_text("a\n", encoding="utf-8")
            digest = hashlib.sha256((root / "nested" / "a.txt").read_bytes()).hexdigest()
            (root / "SHA256SUMS").write_text(f"{digest}  nested/a.txt\n", encoding="utf-8")
            verify.verify_seal(root)
            (root / "extra.txt").write_text("extra\n", encoding="utf-8")
            with self.assertRaises(verify.VerificationError):
                verify.verify_seal(root)

    def test_source_inventory_is_exactly_one_transaction_file(self) -> None:
        bindings = {role: verify._role_binding(verify.ROOT, role)[0] for role in ("control", "candidate")}
        self.assertEqual(
            verify.verify_source_pair(verify.ROOT, bindings),
            [verify.EXPECTED_CHANGED_FILE],
        )

    def test_duplicate_json_key_is_rejected(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "duplicate.json"
            path.write_text('{"a": 1, "a": 2}\n', encoding="utf-8")
            with self.assertRaises(verify.VerificationError):
                verify.read_json(path, "duplicate")

    def test_reused_binding_schema_and_role_are_authenticated(self) -> None:
        source_root = verify.ROOT
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            for relative in (
                "protocol.json",
                "control-binding.json",
                "sources/control.json",
                "prior/candidate-binding.json",
                "prior/candidate-source-binding.json",
                "prior/build.json",
                "prior/build.stdout",
                "prior/build.stderr",
            ):
                destination = root / relative
                destination.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(source_root / relative, destination)
            prior_path = root / "prior/candidate-binding.json"
            prior = json.loads(prior_path.read_text(encoding="utf-8"))
            prior["role"] = "control"
            prior["schema"] = "litchi-0471-role-binding-v1"
            prior_path.write_text(json.dumps(prior), encoding="utf-8")
            control_path = root / "control-binding.json"
            control = json.loads(control_path.read_text(encoding="utf-8"))
            control["reused_binding_sha256"] = hashlib.sha256(prior_path.read_bytes()).hexdigest()
            control_path.write_text(json.dumps(control), encoding="utf-8")
            with self.assertRaisesRegex(verify.VerificationError, "reused binding schema/role"):
                verify._role_binding(root, "control")

    def test_live_verification_authenticates_one_shared_checkout_and_both_binaries(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            tree = root / "tree"
            tree.mkdir()
            for relative, expected in verify.FIXTURES.items():
                destination = tree / relative
                destination.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(Path(__file__).parents[4] / relative, destination)
            manifest = root / "sources/control.json"
            manifest.parent.mkdir()
            manifest.write_text("{}\n", encoding="utf-8")
            bindings = {}
            for role in ("control", "candidate"):
                binary = root / f"{role}.bin"
                binary.write_bytes(role.encode("ascii"))
                digest = hashlib.sha256(binary.read_bytes()).hexdigest()
                bindings[role] = {
                    "binary_path": str(binary),
                    "binary_sha256": digest,
                    "bytes": binary.stat().st_size,
                    "build_path": str(tree),
                    "revision": role,
                    "source_manifest": "sources/control.json",
                }
            answers = iter(["control\n", ""])
            with mock.patch.object(verify.subprocess, "check_output", side_effect=lambda *args, **kwargs: next(answers)):
                verify.verify_live_files(root, bindings, {"build_path": str(tree)})


if __name__ == "__main__":
    unittest.main(verbosity=2)
