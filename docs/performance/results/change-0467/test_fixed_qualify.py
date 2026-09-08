#!/usr/bin/env python3
"""Focused pure tests for the 0467 fixed-checkout qualification helpers.

These tests read retained JSON evidence and use temporary copies for negative
custody cases.  They do not build binaries, invoke the benchmark harness, or
run a profiler.
"""

from __future__ import annotations

import copy
import json
from pathlib import Path
import shutil
import tempfile
import unittest

import fixed_qualify as fixed
import pack_fixed


HERE = Path(__file__).resolve().parent


def _load_json(path: Path) -> dict:
    return json.loads(path.read_text(encoding="utf-8"))


def _records() -> dict[str, dict]:
    """Build the smallest valid identity projection for _check_identity."""

    corpus = {
        "name": "xlsx-dense-wide",
        "shape": "dense-wide",
        "archive_sha256": "c" * 64,
    }
    key = (fixed.XLSX_CASE, fixed.normal.canonical_sha256(corpus))

    def record(revision: str, binary: str, catalog_digest: str) -> dict:
        return {
            "binding": {"revision": revision, "binary_sha256": binary},
            "catalog": {
                "corpus_identity_sha256": "d" * 64,
                "canonical_sha256": catalog_digest,
            },
            "indexed": {key: {"corpus": copy.deepcopy(corpus)}},
        }

    return {
        "a1": record("1" * 40, "a" * 64, "1" * 64),
        "b1": record("2" * 40, "b" * 64, "2" * 64),
        "b2": record("2" * 40, "b" * 64, "3" * 64),
        "a2": record("1" * 40, "a" * 64, "4" * 64),
    }


class FixedProtocolTests(unittest.TestCase):
    def test_frozen_fixed_protocol_has_exact_primary_scope(self):
        protocol = fixed._validate_protocol(HERE / "protocol-fixed.json")
        self.assertEqual(protocol["samples"], 500)
        self.assertEqual(protocol["warmups"], 5)
        self.assertEqual(protocol["order"], list(fixed.FIXED_LANES))
        self.assertEqual(protocol["cases"], list(fixed.FIXED_CASES))
        self.assertEqual(protocol["shapes"], [fixed.XLSX_SHAPE])
        self.assertEqual(protocol["expected_result_count"], 4)

    def test_protocol_change_is_rejected(self):
        protocol = _load_json(HERE / "protocol-fixed.json")
        protocol["expected_result_count"] = 3
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "protocol-fixed.json"
            path.write_text(json.dumps(protocol) + "\n", encoding="utf-8")
            with self.assertRaises(fixed.QualificationError):
                fixed._validate_protocol(path)


class FixedIdentityTests(unittest.TestCase):
    def test_distinct_build_catalog_custody_with_shared_corpus_passes(self):
        records = _records()
        fixed._check_identity(records)
        self.assertNotEqual(
            records["a1"]["catalog"]["canonical_sha256"],
            records["b1"]["catalog"]["canonical_sha256"],
        )
        self.assertEqual(
            {
                record["catalog"]["corpus_identity_sha256"]
                for record in records.values()
            },
            {"d" * 64},
        )

    def test_different_corpus_catalog_projection_is_rejected(self):
        records = _records()
        records["b2"]["catalog"]["corpus_identity_sha256"] = "e" * 64
        with self.assertRaises(fixed.QualificationError):
            fixed._check_identity(records)


class FixedCustodyAndScopeTests(unittest.TestCase):
    def test_primary_projection_keeps_only_claim_scope_row_and_parallel_identity(self):
        report = _load_json(HERE / "A1-fixed/report.json")
        projected = fixed._primary_report_projection(report, "A1-fixed")
        self.assertEqual(len(projected["results"]), 1)
        self.assertEqual(projected["results"][0]["case"], fixed.XLSX_CASE)
        self.assertEqual(
            projected["results"][0]["corpus"]["shape"], fixed.XLSX_SHAPE
        )
        self.assertEqual(projected["configuration"]["cases"], [fixed.XLSX_CASE])
        self.assertEqual(len(projected["parallel_metrics"]["cases"]), 1)
        self.assertEqual(
            projected["parallel_metrics"]["cases"][0]["case"], fixed.XLSX_CASE
        )

    def test_tampered_build_receipt_is_rejected_by_binding_custody(self):
        retained = (
            "control-fixed-binding.json",
            "control-fixed-build.json",
            "control-binding.json",
            "source-bindings.json",
            "sources/control.json",
        )
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            for relative in retained:
                destination = root / relative
                destination.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(HERE / relative, destination)

            build_path = root / "control-fixed-build.json"
            build = _load_json(build_path)
            build["clean_after"] = False
            build_path.write_text(
                json.dumps(build, indent=2, sort_keys=True) + "\n",
                encoding="utf-8",
            )
            with self.assertRaises(fixed.QualificationError):
                fixed._fixed_binding(root, "control")

    def test_out_of_scope_shape_is_rejected(self):
        report = _load_json(HERE / "A1-fixed/report.json")
        binding = {
            "revision": report["environment"]["git_revision"],
            "binary_sha256": report["binary_identity"]["binary_sha256"],
            "bytes": report["binary_identity"]["binary_bytes"],
        }
        changed = copy.deepcopy(report)
        xlsx_row = next(
            row
            for row in changed["results"]
            if row["case"] == fixed.XLSX_CASE
        )
        xlsx_row["corpus"]["shape"] = "medium"
        with self.assertRaises(fixed.QualificationError):
            fixed._fixed_report(changed, "synthetic-out-of-scope", binding)

    def test_claim_proposal_is_xlsx_only_and_keeps_doc_guard_marker(self):
        summary_result = {
            "case": fixed.XLSX_CASE,
            "shape": fixed.XLSX_SHAPE,
            "corpus": {
                "name": "xlsx-dense-wide",
                "archive_sha256": "a" * 64,
                "generator": "litchi-xlsx-synthetic-v1",
                "shape": fixed.XLSX_SHAPE,
                "package_format": "XLSX/OPC/ZIP",
            },
            "elapsed_ns": {
                "accepted_statistics": ["p50", "mean"],
                "adverse_both_statistics": ["p99"],
            },
        }
        proposal = fixed._claim_proposal(summary_result)
        entry = proposal["entry"]
        self.assertTrue(proposal["proposal_only"])
        self.assertEqual(entry["scope"]["selectors"], [fixed.XLSX_CASE])
        self.assertTrue(entry["latency_evidence"]["doc_rows_are_guards_only"])
        self.assertEqual(entry["latency_evidence"]["accepted_cells"], 2)
        self.assertEqual(entry["latency_evidence"]["adverse_both_cells"], 1)


class FixedPackagePlanTests(unittest.TestCase):
    def test_package_plan_retains_all_four_fixed_roles(self):
        self.assertEqual(tuple(pack_fixed.FIXED_REPORTS), ("a1", "b1", "b2", "a2"))
        self.assertEqual(
            tuple(pack_fixed.FIXED_ARTIFACT_NAMES), ("a1", "b1", "b2", "a2")
        )
        self.assertEqual(
            pack_fixed.PACKAGE_MANIFEST_NAME,
            "0467-xlsx-cell-attributes-abba-manifest.json",
        )
        self.assertEqual(pack_fixed.PACKAGE_SUMMARY_NAME, "summary.json")


if __name__ == "__main__":
    unittest.main()
