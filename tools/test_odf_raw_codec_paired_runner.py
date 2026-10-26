#!/usr/bin/env python3
"""Source-only tests for the ODF raw-codec manifest and paired runner."""

from __future__ import annotations

import copy
import json
from pathlib import Path
import tempfile
import unittest

from tools.native_qualification import odf_raw_codec_paired_runner as runner


ROOT = Path(__file__).resolve().parents[1]
MANIFEST_PATH = ROOT / "tools/native_qualification/manifests/odf-raw-codec.json"


def manifest() -> dict:
    return json.loads(MANIFEST_PATH.read_text(encoding="utf-8"))


class OdfRawCodecRunnerTests(unittest.TestCase):
    def test_manifest_binds_all_routes_and_exact_protocol(self) -> None:
        value = runner.validate_manifest(manifest())
        self.assertEqual([route["id"] for route in value["routes"]], list(runner.ROUTE_IDS))
        self.assertEqual(value["paired"]["abba_order"], list(runner.ABBA_ORDER))
        self.assertEqual(value["paired"]["warmup_iterations_per_leg"], runner.WARMUPS)
        self.assertEqual(value["paired"]["samples_per_leg"], runner.SAMPLES)
        self.assertEqual(value["paired"]["probe_iterations"], runner.PROBE_ITERATIONS)
        self.assertIn(runner.NON_ACCEPTANCE, value["claim_scope"])

    def test_manifest_rejects_stale_hashes_and_route_mutation(self) -> None:
        value = manifest()
        value["source"]["candidate_lock_sha256"] = None
        with self.assertRaises(runner.RunnerError):
            runner.validate_manifest(value)
        value = manifest()
        value["routes"][0], value["routes"][1] = value["routes"][1], value["routes"][0]
        with self.assertRaises(runner.RunnerError):
            runner.validate_manifest(value)

    def test_build_journal_is_variant_major_and_hashed(self) -> None:
        events = [{"event_id": f"build-{key.replace('/', '-')}"} for key in runner._build_keys()]
        journal = runner._journal(events)
        self.assertEqual(
            [event["event_id"] for event in journal["events"]],
            ["build-baseline-normal", "build-candidate-normal", "build-baseline-allocator", "build-candidate-allocator"],
        )
        self.assertEqual(runner._journal_events(journal)["build-candidate-allocator"], events[-1])
        broken = copy.deepcopy(journal)
        broken["events"][1], broken["events"][2] = broken["events"][2], broken["events"][1]
        with self.assertRaises(runner.RunnerError):
            runner._journal_events(broken)

    def test_fresh_file_rejects_existing_and_symlinked_ancestors(self) -> None:
        with tempfile.TemporaryDirectory(prefix="odf-raw-runner-") as directory:
            root = Path(directory)
            existing = root / "existing"
            existing.write_text("already", encoding="utf-8")
            instance = runner.PairedRunner.__new__(runner.PairedRunner)
            instance.baseline_root = root / "baseline"
            instance.candidate_root = root / "candidate"
            with self.assertRaises(runner.RunnerError):
                instance._fresh_file(existing, "existing")
            target = root / "target"
            target.mkdir()
            link = root / "link"
            link.symlink_to(target, target_is_directory=True)
            with self.assertRaises(runner.RunnerError):
                instance._fresh_file(link / "result.json", "symlinked output")

    def test_probe_source_contains_all_negative_and_fidelity_proofs(self) -> None:
        probe = (ROOT / "tools/perf-baseline/examples/odf_raw_codec_qualification.rs").read_text(encoding="utf-8")
        for route in runner.ROUTE_IDS:
            self.assertIn(route, probe)
        for case in manifest()["phases"]["negative_cases"]:
            self.assertIn(case, probe)
        for proof in ("source_bytes_ok", "facade_bytes_ok", "untouched_members_ok", "negative_cases_ok"):
            self.assertIn(f'"{proof}"', probe)
        self.assertIn("PackageCompression::Stored", probe)


if __name__ == "__main__":
    unittest.main()
