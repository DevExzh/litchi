#!/usr/bin/env python3
"""Mutation tests for the independent 0475 evidence verifier.

These tests exercise the verifier's rejection boundaries with small temporary
fixtures.  They do not rerun a profile or decompress the multi-gigabyte raw
Heaptrack streams.
"""

from __future__ import annotations

import copy
import gzip
import hashlib
import json
from pathlib import Path
import shutil
import tempfile
import unittest

import verify


ROOT = Path(__file__).resolve().parent


class EvidenceMutationTests(unittest.TestCase):
    def test_protocol_selector_mutation_is_rejected(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            shutil.copy2(ROOT / "capture.py", root / "capture.py")
            protocol = json.loads((ROOT / "protocol.json").read_text(encoding="utf-8"))
            protocol["selector"] = "stale_case"
            (root / "protocol.json").write_text(json.dumps(protocol), encoding="utf-8")
            with self.assertRaises(verify.VerificationError):
                verify.verify_protocol(root)

    def test_elapsed_mean_mutation_is_rejected(self) -> None:
        report = json.loads((ROOT / "normal-R1/report.json").read_text(encoding="utf-8"))
        elapsed = report["results"][0]["elapsed_ns"]
        elapsed["mean"] += 1
        with self.assertRaises(verify.VerificationError):
            verify.verify_elapsed(elapsed, 30, "mutated elapsed")

    def test_sink_counter_mutation_changes_producer_projection(self) -> None:
        current = json.loads((ROOT / "normal-R1/report.json").read_text(encoding="utf-8"))
        row = current["results"][0]
        expected = verify.expected_producer_projection(row, 30)
        mutated = copy.deepcopy(row)
        mutated["sink"]["accepted_bytes"] += 1
        self.assertNotEqual(expected, verify.expected_producer_projection(mutated, 30))

    def test_capture_receipt_hash_mutation_is_rejected(self) -> None:
        binding = json.loads((ROOT / "binding.json").read_text(encoding="utf-8"))
        protocol_hash = verify.sha256(ROOT / "protocol.json")
        binding_hash = verify.sha256(ROOT / "binding.json")
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            lane = root / "normal-R1"
            shutil.copytree(ROOT / "normal-R1", lane)
            receipt_path = lane / "receipt.json"
            receipt = json.loads(receipt_path.read_text(encoding="utf-8"))
            receipt["artifacts"]["report.json"]["sha256"] = "0" * 64
            receipt_path.write_text(json.dumps(receipt), encoding="utf-8")
            with self.assertRaises(verify.VerificationError):
                verify.verify_capture_receipt(root, ("normal-R1", "normal", 30, 3), binding, binding_hash, protocol_hash, {}, {})

    def test_compressed_payload_mutation_is_rejected(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            lane = root / "cpu-P1"
            lane.mkdir()
            compressed = lane / "perf.data.gz"
            with compressed.open("wb") as output:
                with gzip.GzipFile(filename="", fileobj=output, mode="wb", mtime=0) as stream:
                    stream.write(b"portable profile")
            item = {
                "path": "cpu-P1/perf.data",
                "sha256": hashlib.sha256(b"portable profile").hexdigest(),
                "bytes": len(b"portable profile"),
                "compressed_path": "cpu-P1/perf.data.gz",
                "compressed_sha256": verify.sha256(compressed),
                "compressed_bytes": compressed.stat().st_size,
            }
            (root / "compression.json").write_text(json.dumps({"schema": verify.COMPRESSION_SCHEMA, "artifacts": [item]}), encoding="utf-8")
            verify._compression_records(root, root / "compression.json", frozenset({"cpu-P1/perf.data"}), "test compression")
            compressed.write_bytes(compressed.read_bytes() + b"mutation")
            with self.assertRaises(verify.VerificationError):
                verify._compression_records(root, root / "compression.json", frozenset({"cpu-P1/perf.data"}), "test compression")

    def test_unsupported_counter_status_is_retained_and_unknown_event_rejected(self) -> None:
        text = (ROOT / "counters/counters.csv").read_text(encoding="utf-8")
        counters = verify.parse_counter_text(text)
        self.assertEqual(counters["LLC-load-misses:u"]["status"], "not_supported")
        self.assertIsNone(counters["LLC-load-misses:u"]["count"])
        with self.assertRaises(verify.VerificationError):
            verify.parse_counter_text(text.replace("context-switches", "future-event"))


if __name__ == "__main__":
    unittest.main()
