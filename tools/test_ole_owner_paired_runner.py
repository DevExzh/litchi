#!/usr/bin/env python3
"""Fail-closed contract tests for the OLE owner paired runner."""

from __future__ import annotations

import copy
import json
from pathlib import Path
import tempfile
import unittest
from unittest import mock

from tools.native_qualification import ole_owner_paired_runner as runner


ROOT = Path(__file__).resolve().parents[1]
MANIFEST = ROOT / "tools/native_qualification/manifests/ole-owner.json"


class OleOwnerPairedRunnerTests(unittest.TestCase):
    def test_toolchain_ignores_ambient_overrides_and_binds_child_environment(self) -> None:
        with tempfile.TemporaryDirectory(prefix="ole-owner-toolchain-") as directory:
            root = Path(directory)
            path_bin = root / "path-bin"
            path_bin.mkdir()
            override_bin = root / "override-bin"
            override_bin.mkdir()

            def executable(path: Path, output: str) -> None:
                path.write_text(f"#!/bin/sh\nprintf '%s\\n' '{output}'\n", encoding="utf-8")
                path.chmod(0o755)

            cargo = path_bin / "cargo"
            rustc = path_bin / "rustc"
            override_cargo = override_bin / "cargo-override"
            override_rustc = override_bin / "rustc-override"
            executable(cargo, "PATH-CARGO")
            executable(rustc, "PATH-RUSTC")
            executable(override_cargo, "OVERRIDE-CARGO")
            executable(override_rustc, "OVERRIDE-RUSTC")

            probe = object.__new__(runner.PairedRunner)
            probe.driver_roots = {"before": root}
            env = {
                "PATH": str(path_bin),
                "CARGO": str(override_cargo),
                "RUSTC": str(override_rustc),
                "RUSTC_WRAPPER": str(override_rustc),
                "RUSTC_WORKSPACE_WRAPPER": str(override_rustc),
                "CARGO_BUILD_RUSTC": str(override_rustc),
                "RUSTUP_TOOLCHAIN": "override-toolchain",
            }
            toolchain = probe._toolchain(env)
            self.assertEqual(toolchain["cargo"]["path"], str(cargo))
            self.assertEqual(toolchain["rustc"]["path"], str(rustc))
            self.assertEqual(toolchain["cargo"]["version"].strip(), "PATH-CARGO")
            self.assertEqual(toolchain["rustc"]["version"].strip(), "PATH-RUSTC")
            self.assertEqual(toolchain["path_env"], str(path_bin))

            probe.toolchain = toolchain
            probe.target_root = root / "targets"
            ambient_after_selection = {**env, "PATH": str(override_bin)}
            with mock.patch.dict(runner.os.environ, ambient_after_selection, clear=True):
                child = probe._environment("before", "normal")
            self.assertEqual(child["CARGO"], str(cargo))
            self.assertEqual(child["RUSTC"], str(rustc))
            self.assertEqual(child["PATH"], str(path_bin))
            for key in ("RUSTC_WRAPPER", "RUSTC_WORKSPACE_WRAPPER", "CARGO_BUILD_RUSTC", "RUSTUP_TOOLCHAIN"):
                self.assertNotIn(key, child)

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

    def test_oracle_allocator_parser_enforces_byte_conservation_and_retained_history(self) -> None:
        header = "\t".join(runner.ALLOCATOR_EVIDENCE_FIELDS)
        rows = [
            f"{route}\t10\t100\t2\t20\t1\t5\t6\t100\t1000\t108\t1081\t110\t1100\t{('129\t128' if route in {'doc_history', 'xls_history'} else '0\t0')}"
            for route in runner.ALLOCATOR_ROUTES
        ]
        with tempfile.TemporaryDirectory(prefix="ole-owner-output-") as directory:
            path = Path(directory) / "allocator.tsv"
            path.write_text("\n".join([header, *rows]) + "\n", encoding="utf-8")
            parsed = runner.PairedRunner._parse_oracle_allocator(path)
            self.assertEqual(len(parsed), len(runner.ALLOCATOR_ROUTES))
            broken = rows.copy()
            broken[0] = broken[0].replace("\t1081\t", "\t1080\t")
            path.write_text("\n".join([header, *broken]) + "\n", encoding="utf-8")
            with self.assertRaisesRegex(runner.RunnerError, "conservation"):
                runner.PairedRunner._parse_oracle_allocator(path)


if __name__ == "__main__":
    unittest.main()
