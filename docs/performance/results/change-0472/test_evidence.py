#!/usr/bin/env python3
"""Focused tests for the 0472 evidence boundary and portable seal."""

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
    def test_chronology_requires_role_build_before_capture(self) -> None:
        def interval(start, finish):
            return {"started_utc": f"2026-09-08T00:00:{start:02d}+00:00",
                    "finished_utc": f"2026-09-08T00:00:{finish:02d}+00:00"}
        builds = {"control": interval(0, 1), "candidate": interval(4, 5)}
        order = ("A1", "B1", "B2", "A2", "A-full", "B-full", "A-heap", "B-heap")
        captures = {lane: interval(2 if i == 0 else 4 + i * 2, 3 if i == 0 else 5 + i * 2)
                    for i, lane in enumerate(order)}
        gates = {"checks": [dict(name="test", **interval(21, 22))]}
        result = verify.verify_chronology(verify.ROOT, builds, captures, gates)
        self.assertTrue(result["candidate_build_after_A1"])
        builds["candidate"] = interval(23, 24)
        with self.assertRaisesRegex(verify.VerificationError, "frozen build/capture chronology"):
            verify.verify_chronology(verify.ROOT, builds, captures, gates)

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

    def test_source_inventory_is_bounded_snapshot_wire_change_set(self) -> None:
        bindings = {role: verify._role_binding(verify.ROOT, role)[0] for role in ("control", "candidate")}
        self.assertEqual(
            verify.verify_source_pair(verify.ROOT, bindings),
            sorted(verify.EXPECTED_CANDIDATE_CHANGED_FILES),
        )

    def test_duplicate_json_key_is_rejected(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "duplicate.json"
            path.write_text('{"a": 1, "a": 2}\n', encoding="utf-8")
            with self.assertRaises(verify.VerificationError):
                verify.read_json(path, "duplicate")

    def test_fresh_role_binding_rejects_prior_chain_fields(self) -> None:
        source_root = verify.ROOT
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            for relative in (
                "protocol.json",
                "control-binding.json",
                "sources/control.json",
                "control-source-binding.json",
                "control-build.json",
                "control-build.stdout",
                "control-build.stderr",
                "control-build.started.json",
            ):
                destination = root / relative
                destination.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(source_root / relative, destination)
            control_path = root / "control-binding.json"
            control = json.loads(control_path.read_text(encoding="utf-8"))
            control["prior_binding_path"] = "control-binding.json"
            control_path.write_text(json.dumps(control), encoding="utf-8")
            with self.assertRaisesRegex(verify.VerificationError, "fresh 0472 binding"):
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
