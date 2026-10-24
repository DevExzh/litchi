#!/usr/bin/env python3
"""Fail-closed contract tests for the OLE owner paired runner."""

from __future__ import annotations

import copy
import json
from pathlib import Path
import tempfile
import unittest

from tools.native_qualification import ole_owner_paired_runner as runner


ROOT = Path(__file__).resolve().parents[1]
MANIFEST = ROOT / "tools/native_qualification/manifests/ole-owner.json"


class OleOwnerPairedRunnerTests(unittest.TestCase):
    def test_manifest_binds_all_routes_and_fixed_resource_policy(self) -> None:
        document, digest = runner._load_manifest(MANIFEST)
        self.assertEqual(len(digest), 64)
        self.assertEqual(document["routes"], list(runner.ROUTES))
        self.assertEqual(document["paired"]["abba_order"], list(runner.ABBA_ORDER))
        self.assertEqual(document["allocator"]["routes"], list(runner.ALLOCATOR_ROUTES))

        for mutation, message in (
            (lambda value: value["paired"].update(blocks=4), "block count"),
            (lambda value: value["quiet"].update(memory_bytes=1), "resource limits"),
            (lambda value: value["allocator"].update(elapsed_claimable=True), "non-claimable"),
            (lambda value: value["routes"].pop(), "route set/order"),
        ):
            broken = copy.deepcopy(document)
            mutation(broken)
            with self.subTest(message=message), tempfile.TemporaryDirectory(prefix="ole-owner-manifest-") as directory:
                path = Path(directory) / "manifest.json"
                path.write_text(json.dumps(broken), encoding="utf-8")
                with self.assertRaisesRegex(runner.RunnerError, message):
                    runner._load_manifest(path)

    def test_allocator_parser_requires_retained_counter_rows(self) -> None:
        probe = object.__new__(runner.PairedRunner)
        header = (
            "w5_alloc,phase,route,iteration,delta_calls,delta_bytes,live_calls,live_bytes,"
            "peak_calls,peak_bytes,peak_delta_calls,peak_delta_bytes,base_live_calls,"
            "base_live_bytes,base_peak_calls,base_peak_bytes"
        )
        row = "w5_alloc,full,doc_history,0," + ",".join("1" for _ in runner.ALLOCATOR_FIELDS)
        with tempfile.TemporaryDirectory(prefix="ole-owner-output-") as directory:
            path = Path(directory) / "allocator.csv"
            path.write_text(
                "route,sample,nanoseconds,output_bytes\n"
                f"{header}\n{row}\n"
                "doc_history,0,20,100\n",
                encoding="utf-8",
            )
            report = probe._parse_allocator(path, "doc_history")
            self.assertEqual(report["rows"][0]["live_bytes"], 1)
            self.assertEqual(report["output_bytes"], 100)

            broken = path.read_text(encoding="utf-8").replace(",1,1,1,1,1,1,1,1,1,1,1,1\n", ",1,1,1,1,1,1,1,1,1,1,1,1\n")
            path.write_text(broken.replace("w5_alloc,full", "w5_alloc,unknown"), encoding="utf-8")
            with self.assertRaisesRegex(runner.RunnerError, "allocator phase rows"):
                probe._parse_allocator(path, "doc_history")

    def test_clone_calibration_accepts_one_sample_per_size_but_not_missing_size(self) -> None:
        probe = object.__new__(runner.PairedRunner)
        with tempfile.TemporaryDirectory(prefix="ole-owner-output-") as directory:
            path = Path(directory) / "clone.csv"
            path.write_text(
                "route,size,sample,nanoseconds,batch,archive_bytes\n"
                "clone_drop,small,0,500000000,1000,128\n"
                "clone_drop,large,0,500000000,1000,4096\n",
                encoding="utf-8",
            )
            rows = probe._parse_clone(path, 2, samples=1)
            self.assertEqual(len(rows), 2)
            path.write_text(path.read_text(encoding="utf-8").replace("clone_drop,large", "clone_drop,small"), encoding="utf-8")
            with self.assertRaisesRegex(runner.RunnerError, "clone small rows"):
                probe._parse_clone(path, 2, samples=1)


if __name__ == "__main__":
    unittest.main()
