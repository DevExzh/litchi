#!/usr/bin/env python3
"""Focused fail-closed tests for the 0468 bundle verifier."""

from __future__ import annotations

import gzip
import hashlib
import json
from pathlib import Path
import tempfile
import unittest
from unittest import mock

import verify


class VerifierHelperTests(unittest.TestCase):
    def test_safe_relative_paths_reject_traversal(self):
        with self.assertRaises(verify.VerificationError):
            verify.relative("../outside", "path")
        with self.assertRaises(verify.VerificationError):
            verify.relative("/absolute", "path")

    def test_json_nonfinite_value_is_rejected(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "bad.json"
            path.write_text('{"value": NaN}\n', encoding="utf-8")
            with self.assertRaises(verify.VerificationError):
                verify.read_json(path, "bad.json")

    def test_compression_hashes_both_gzip_and_decompressed_bytes(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            target = root / "samples-fp"
            target.mkdir()
            plain = b"cycles:u\nprofile bytes\n"
            compressed = target / "perf.data.gz"
            with gzip.open(compressed, "wb") as stream:
                stream.write(plain)
            script = target / "perf-script.stdout.gz"
            with gzip.open(script, "wb") as stream:
                stream.write(plain + b"second sample\n")
            rows = []
            for plain_name, compressed_path in (
                ("samples-fp/perf.data", compressed),
                ("samples-fp/perf-script.stdout", script),
            ):
                decompressed = gzip.open(compressed_path, "rb").read()
                rows.append({
                    "path": plain_name,
                    "sha256": hashlib.sha256(decompressed).hexdigest(),
                    "bytes": len(decompressed),
                    "compressed_path": plain_name + ".gz",
                    "compressed_sha256": hashlib.sha256(compressed_path.read_bytes()).hexdigest(),
                    "compressed_bytes": compressed_path.stat().st_size,
                })
            (root / "compression.json").write_text(json.dumps({
                "schema": "litchi-0468-compression-v1", "artifacts": rows,
            }), encoding="utf-8")
            self.assertEqual(verify._verify_compression(root)["artifacts"], sorted({
                "samples-fp/perf.data", "samples-fp/perf-script.stdout",
            }))
            rows[0]["compressed_sha256"] = "0" * 64
            (root / "compression.json").write_text(json.dumps({
                "schema": "litchi-0468-compression-v1", "artifacts": rows,
            }), encoding="utf-8")
            with self.assertRaises(verify.VerificationError):
                verify._verify_compression(root)

    def test_sha256sums_requires_exact_file_coverage(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            member = root / "receipt.json"
            member.write_text("{}\n", encoding="utf-8")
            checksum = hashlib.sha256(member.read_bytes()).hexdigest()
            (root / "SHA256SUMS").write_text(f"{checksum}  receipt.json\n", encoding="utf-8")
            self.assertEqual(verify._verify_sha256sums(root), 1)
            (root / "unlisted.bin").write_bytes(b"tamper")
            with self.assertRaises(verify.VerificationError):
                verify._verify_sha256sums(root)

    def test_elapsed_percentile_tamper_is_rejected_by_shared_comparator(self):
        comparator, _ = verify._load_perf_compare(Path(verify.ROOT).parents[3])
        row = {
            "elapsed_ns": {
                "unit": "ns",
                "samples": [1, 2, 3, 4, 5],
                "p50": 3,
                "p95": 5,
                "p99": 5,
            }
        }
        comparator._latencies(row, "row", 5)
        row["elapsed_ns"]["p50"] = 99
        with self.assertRaises(ValueError):
            comparator._latencies(row, "row", 5)

    def test_additional_context_tamper_is_rejected(self):
        expected = {"schema": "litchi-0468-profile-context-v1", "rows": {"parser": 4}}

        class Analyzer:
            @staticmethod
            def profile_context(_script, _reports, *, top):
                assert top == 40
                return expected

        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            target = root / "additional-summary.json"
            target.write_text(json.dumps(expected), encoding="utf-8")
            self.assertEqual(
                verify._verify_additional_context(root, Analyzer, root / "script.gz", []),
                {"path": "additional-summary.json"},
            )
            target.write_text(json.dumps({**expected, "rows": {"parser": 5}}), encoding="utf-8")
            with self.assertRaises(verify.VerificationError):
                verify._verify_additional_context(root, Analyzer, root / "script.gz", [])

    def test_preseal_live_binary_hash_mismatch_is_rejected(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            tree = root / "tree"
            tree.mkdir()
            binary = root / "profile"
            binary.write_bytes(b"authenticated profile")
            expected_sha = hashlib.sha256(binary.read_bytes()).hexdigest()
            binding = {
                "clean_tree": str(tree),
                "binary_path": str(binary),
                "binary_sha256": expected_sha,
                "bytes": binary.stat().st_size,
                "included_fixtures": {},
            }
            (root / "sources.json").write_text("{}\n", encoding="utf-8")

            def git_output(argv, **_kwargs):
                return verify.REVISION + "\n" if argv[-2:] == ["rev-parse", "HEAD"] else ""

            with mock.patch.object(verify.subprocess, "check_output", side_effect=git_output):
                verify._verify_live_state(root, binding, preseal=True)
                tampered = {**binding, "binary_sha256": "0" * 64}
                with self.assertRaises(verify.VerificationError):
                    verify._verify_live_state(root, tampered, preseal=True)


if __name__ == "__main__":
    raise SystemExit(unittest.main())
