"""Focused tests for the release performance report/catalog binding gate."""

from __future__ import annotations

import copy
import hashlib
import json
import tempfile
import unittest
from pathlib import Path

from tools.generate_corpus_manifest_v2 import canonical_bytes, generate
from tools.validate_perf_corpus_binding import (
    ValidationError,
    validate_binding,
    validate_paths,
)


ROOT = Path(__file__).resolve().parents[1]
V1_PATH = ROOT / "docs/performance/results/perf-regression-default-manifest-v1.json"
V2_PATH = ROOT / "docs/performance/results/perf-corpus-manifest-v2.json"


def _refresh_hashes(catalog: dict) -> None:
    content_set = {
        "corpora": [
            {
                "id": corpus["id"],
                "archive_sha256": corpus["bytes"]["archive_sha256"],
                "members": [
                    {
                        "ordinal": member["ordinal"],
                        "name": member["name"],
                        "sha256": member["sha256"],
                    }
                    for member in corpus["members"]["items"]
                ],
            }
            for corpus in catalog["corpora"]
        ],
        "case_bindings": [
            {
                "case": binding["case"],
                "corpus_id": binding["corpus_id"],
                "role": binding["role"],
            }
            for binding in catalog["case_bindings"]
        ],
    }
    catalog["content_set_sha256"] = hashlib.sha256(
        canonical_bytes(content_set)
    ).hexdigest()
    without_hash = dict(catalog)
    without_hash.pop("catalog_sha256")
    catalog["catalog_sha256"] = hashlib.sha256(
        canonical_bytes(without_hash)
    ).hexdigest()


class PerfCorpusBindingTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.v1 = json.loads(V1_PATH.read_text(encoding="utf-8"))
        cls.case = next(iter(cls.v1["case_corpora"]))
        cls.corpus_name = cls.v1["case_corpora"][cls.case][0]
        cls.corpus = cls.v1["corpora"][cls.corpus_name]
        cls.catalog = generate(
            {
                **cls.v1,
                "case_corpora": {cls.case: [cls.corpus_name]},
            },
            revision="test-revision",
        )
        cls.report = {
            "schema_version": 1,
            "corpus_catalog": {
                "manifest_version": cls.catalog["manifest_version"],
                "catalog_id": cls.catalog["catalog_id"],
                "catalog_sha256": cls.catalog["catalog_sha256"],
                "content_set_sha256": cls.catalog["content_set_sha256"],
            },
            "results": [{"case": cls.case, "corpus": cls.corpus}],
        }

    def test_valid_report_and_catalog_bind(self) -> None:
        self.assertEqual(
            validate_binding(self.report, self.catalog),
            (len(self.catalog["corpora"]), 1),
        )

    def test_catalog_hash_tampering_is_rejected(self) -> None:
        catalog = copy.deepcopy(self.catalog)
        catalog["corpora"][0]["name"] = "tampered"
        with self.assertRaises(ValidationError):
            validate_binding(self.report, catalog)

    def test_report_reference_tampering_is_rejected(self) -> None:
        report = copy.deepcopy(self.report)
        report["corpus_catalog"]["content_set_sha256"] = "0" * 64
        with self.assertRaises(ValidationError):
            validate_binding(report, self.catalog)

    def test_recomputed_semantic_hash_tampering_is_rejected(self) -> None:
        catalog = copy.deepcopy(self.catalog)
        original = catalog["corpora"][0]["bytes"]["archive_sha256"]
        catalog["corpora"][0]["bytes"]["archive_sha256"] = (
            "0" * 64 if original != "0" * 64 else "1" * 64
        )
        _refresh_hashes(catalog)
        report = copy.deepcopy(self.report)
        report["corpus_catalog"].update(
            {
                "catalog_sha256": catalog["catalog_sha256"],
                "content_set_sha256": catalog["content_set_sha256"],
            }
        )
        with self.assertRaises(ValidationError):
            validate_binding(report, catalog)

    def test_duplicate_json_keys_are_rejected_by_validate_paths(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            report_path = Path(directory) / "report.json"
            catalog_path = Path(directory) / "catalog.json"
            report_path.write_text(json.dumps(self.report), encoding="utf-8")
            catalog_json = json.dumps(self.catalog)
            catalog_path.write_text(
                catalog_json[:-1]
                + ',"catalog_sha256":"'
                + self.catalog["catalog_sha256"]
                + '"}',
                encoding="utf-8",
            )
            with self.assertRaises(ValidationError):
                validate_paths(report_path, catalog_path)

    def test_list_valued_report_case_is_rejected_as_validation_error(self) -> None:
        report = copy.deepcopy(self.report)
        report["results"][0]["case"] = [self.case]
        with self.assertRaises(ValidationError):
            validate_binding(report, self.catalog)

    def test_missing_nested_field_is_rejected_as_validation_error(self) -> None:
        catalog = copy.deepcopy(self.catalog)
        del catalog["corpora"][0]["bytes"]["archive_sha256"]
        with self.assertRaises(ValidationError):
            validate_binding(self.report, catalog)

    def test_binding_metadata_mismatch_is_rejected_after_rehashing(self) -> None:
        catalog = copy.deepcopy(self.catalog)
        catalog["case_bindings"][0]["legacy_name"] = "not-the-referenced-corpus"
        _refresh_hashes(catalog)
        report = copy.deepcopy(self.report)
        report["corpus_catalog"].update(
            {
                "catalog_sha256": catalog["catalog_sha256"],
                "content_set_sha256": catalog["content_set_sha256"],
            }
        )
        with self.assertRaises(ValidationError):
            validate_binding(report, catalog)

    def test_checked_rust_catalog_binds_to_report_reconstructed_from_bindings(self) -> None:
        catalog = json.loads(V2_PATH.read_text(encoding="utf-8"))
        corpora_by_id = {corpus["id"]: corpus for corpus in catalog["corpora"]}
        report = {
            "schema_version": 1,
            "corpus_catalog": {
                "manifest_version": catalog["manifest_version"],
                "catalog_id": catalog["catalog_id"],
                "catalog_sha256": catalog["catalog_sha256"],
                "content_set_sha256": catalog["content_set_sha256"],
            },
            "results": [
                {
                    "case": binding["case"],
                    "corpus": corpora_by_id[binding["corpus_id"]]["legacy_v1"],
                }
                for binding in catalog["case_bindings"]
            ],
        }
        self.assertEqual(
            validate_binding(report, catalog),
            (len(catalog["corpora"]), len(catalog["case_bindings"])),
        )


if __name__ == "__main__":
    unittest.main()
