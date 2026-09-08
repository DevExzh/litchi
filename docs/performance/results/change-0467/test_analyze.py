#!/usr/bin/env python3
"""Focused, non-profiler tests for the 0467 ABBA analyzers."""

from __future__ import annotations

import copy
import json
from pathlib import Path
import shutil
import tempfile
import unittest

import analyze
import qualify


HERE = Path(__file__).resolve().parent


def _write_json(path: Path, value: object) -> None:
    path.write_text(json.dumps(value, indent=2) + "\n", encoding="utf-8")


class PureMetricTests(unittest.TestCase):
    def test_ratio_and_drift_use_candidate_over_control_and_five_percent_trigger(self):
        ratio = analyze._ratio_percent(100.0, 94.0, "ratio")
        self.assertEqual(ratio["ratio_current_over_baseline"], 0.94)
        self.assertAlmostEqual(ratio["reduction_percent"], 6.0)
        self.assertFalse(ratio["regression_over_5_percent"])
        drift = analyze._drift(100.0, 106.0, "drift")
        self.assertTrue(drift["review_triggered"])
        self.assertAlmostEqual(drift["delta_percent"], 6.0)

    def test_rss_parser_requires_a_positive_gnu_time_value(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "resource.log"
            path.write_text(
                "Maximum resident set size (kbytes): 12345\n", encoding="utf-8"
            )
            parsed = analyze._resource_rss(path)
            self.assertEqual(parsed["max_rss_kib"], 12345)
            path.write_text("no rss\n", encoding="utf-8")
            with self.assertRaises(analyze.AnalysisError):
                analyze._resource_rss(path)


class StrictQualificationPureTests(unittest.TestCase):
    def test_frozen_500_protocol_and_claim_entry_shape(self):
        protocol = qualify._validate_protocol(HERE / "protocol-500.json")
        self.assertEqual(protocol["samples"], qualify.SAMPLES)
        self.assertEqual(protocol["order"], list(qualify.STRICT_LANES))
        summary_result = {
            "case": qualify.PRIMARY_CASE,
            "shape": qualify.PRIMARY_SHAPE,
            "corpus": {
                "name": "xlsx-dense-wide",
                "archive_sha256": "a" * 64,
                "generator": "litchi-xlsx-synthetic-v1",
                "shape": "dense-wide",
                "package_format": "XLSX/OPC/ZIP",
            },
            "elapsed_ns": {
                "accepted_statistics": ["p50", "mean", "p95", "p99"],
                "adverse_both_statistics": [],
            },
        }
        proposal = qualify._proposal(summary_result, strict_summary_path="summary.json")
        entry = proposal["entry"]
        self.assertTrue(proposal["proposal_only"])
        self.assertEqual(entry["change_id"], qualify.PACKAGE_CHANGE_ID)
        self.assertEqual(entry["latency_evidence"]["accepted_cells"], 4)
        self.assertEqual(entry["latency_evidence"]["adverse_both_cells"], 0)
        self.assertEqual(proposal["package_manifest_name"], qualify.PACKAGE_MANIFEST_NAME)


class SyntheticFormalBundleTests(unittest.TestCase):
    """Exercise all identity gates from a copied report, without a build."""

    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)
        source = HERE / "A1"
        for name in ("protocol-r1.json",):
            shutil.copy2(HERE / name, self.root / name)
        shutil.copy2(source / "resource.log", self.root / "resource.log")

        control_report = json.loads((source / "report.json").read_text())
        source_catalog = json.loads((source / "corpus-catalog.json").read_text())
        control_report["environment"]["git_worktree_dirty"] = False
        control_report["environment"]["cpu_affinity"] = "2"
        control_report["binary_identity"]["path"] = "/tmp/formal-control"
        control_report["binary_identity"]["binary_sha256"] = "a" * 64
        control_report["binary_identity"]["binary_bytes"] = 100
        self.control_revision = "1" * 40
        self.candidate_revision = "2" * 40
        self.candidate_sha = "b" * 64

        control_binding = {
            "revision": self.control_revision,
            "binary_sha256": "a" * 64,
            "bytes": 100,
            "build_receipt": "control-build.json",
            "build_receipt_sha256": "",
            "clean_build": True,
        }
        candidate_binding = {
            "revision": self.candidate_revision,
            "binary_sha256": self.candidate_sha,
            "bytes": 100,
            "build_receipt": "candidate-build.json",
            "build_receipt_sha256": "",
            "clean_build": True,
        }
        for role, binding in (("control", control_binding), ("candidate", candidate_binding)):
            build = {
                "role": role,
                "revision": binding["revision"],
                "argv": ["cargo", "build", "--release"],
                "clean_before": True,
                "clean_after": True,
                "exit_code": 0,
            }
            build_path = self.root / binding["build_receipt"]
            _write_json(build_path, build)
            binding["build_receipt_sha256"] = analyze.raw_sha256(build_path)
        _write_json(self.root / "control-binding.json", control_binding)
        _write_json(self.root / "candidate-binding.json", candidate_binding)

        for lane in analyze.LANE_NAMES:
            lane_dir = self.root / lane
            lane_dir.mkdir()
            role = analyze.LANE_ROLES[lane]
            is_candidate = role == "candidate"
            report = copy.deepcopy(control_report)
            report["environment"]["git_revision"] = (
                self.candidate_revision if is_candidate else self.control_revision
            )
            report["binary_identity"]["binary_sha256"] = (
                self.candidate_sha if is_candidate else "a" * 64
            )
            report["binary_identity"]["path"] = (
                "/tmp/formal-candidate" if is_candidate else "/tmp/formal-control"
            )
            report["binary_identity"]["binary_bytes"] = 100
            catalog = copy.deepcopy(source_catalog)
            catalog["build"]["git_revision"] = report["environment"]["git_revision"]
            catalog["build"]["git_worktree_dirty"] = False
            catalog_without_hash = {
                key: value for key, value in catalog.items() if key != "catalog_sha256"
            }
            catalog["catalog_sha256"] = analyze.canonical_sha256(catalog_without_hash)
            report["corpus_catalog"]["catalog_sha256"] = catalog["catalog_sha256"]
            _write_json(lane_dir / "report.json", report)
            _write_json(lane_dir / "corpus-catalog.json", catalog)
            shutil.copy2(source / "resource.log", lane_dir / "resource.log")
            receipt = {
                "schema": "litchi-0467-capture-v1",
                "lane": lane,
                "role": role,
                "revision": report["environment"]["git_revision"],
                "binary_sha256": report["binary_identity"]["binary_sha256"],
                "binding_sha256": analyze.raw_sha256(
                    self.root / f"{role}-binding.json"
                ),
                "clean_before": True,
                "clean_after": True,
                "binary_unchanged": True,
                "report_metadata_matches_clean_role": True,
                "samples": analyze.SAMPLES,
                "warmups": analyze.WARMUPS,
                "argv": ["taskset", "-c", "2", "program", "--workers", "1"],
                "exit_code": 0,
            }
            _write_json(lane_dir / "receipt.json", receipt)

    def tearDown(self):
        self.temp.cleanup()

    def test_formal_analysis_records_rows_and_canonical_summary(self):
        result = analyze.analyze(self.root)
        self.assertEqual(result["schema"], analyze.SCHEMA)
        self.assertEqual(len(result["results"]), 6)
        self.assertTrue(result["verification"]["clean_worktree_verified"])
        self.assertTrue(result["verification"]["canonical_abba_summary_verified"])
        self.assertEqual(result["abba_summary"]["verification"]["result_count"], 6)
        self.assertFalse(result["claim_registration"]["strict_latency_claim_eligible"])
        self.assertEqual(result["claim_registration"]["registry_minimum_samples_per_case"], 500)
        for row in result["results"]:
            self.assertEqual(set(row["statistics_ns"]["a1"]), set(analyze.STATISTICS))
            self.assertEqual(row["review_triggers"]["threshold_percent"], 5.0)
            self.assertEqual(row["process_rss_kib"]["a1"], row["process_rss_kib"]["b1"])

    def test_catalog_build_identity_differs_while_corpus_identity_matches(self):
        catalogs = [
            analyze._catalog_file_identity(
                self.root,
                self.root / lane,
                json.loads((self.root / lane / "report.json").read_text()),
                lane,
            )
            for lane in analyze.LANE_NAMES
        ]
        raw_catalogs = [
            json.loads((self.root / lane / "corpus-catalog.json").read_text())
            for lane in analyze.LANE_NAMES
        ]
        self.assertNotEqual(raw_catalogs[0]["build"]["git_revision"], raw_catalogs[2]["build"]["git_revision"])
        self.assertNotEqual(catalogs[0]["canonical_sha256"], catalogs[2]["canonical_sha256"])
        self.assertNotEqual(catalogs[0]["catalog_sha256"], catalogs[2]["catalog_sha256"])
        self.assertEqual(
            {catalog["corpus_identity_sha256"] for catalog in catalogs},
            {catalogs[0]["corpus_identity_sha256"]},
        )

    def test_dirty_report_is_rejected_even_when_other_identity_is_valid(self):
        report_path = self.root / "A1-clean" / "report.json"
        report = json.loads(report_path.read_text())
        report["environment"]["git_worktree_dirty"] = True
        _write_json(report_path, report)
        with self.assertRaises(analyze.AnalysisError):
            analyze.analyze(self.root)

    def test_candidate_revision_or_binary_mismatch_is_rejected(self):
        binding_path = self.root / "candidate-binding.json"
        binding = json.loads(binding_path.read_text())
        binding["revision"] = self.control_revision
        _write_json(binding_path, binding)
        with self.assertRaises(analyze.AnalysisError):
            analyze.analyze(self.root)


if __name__ == "__main__":
    unittest.main()
