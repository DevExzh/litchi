"""Adversarial tests for the non-iWork CRUD coverage index gate."""

from __future__ import annotations

import copy
import hashlib
import json
import tempfile
import unittest
from pathlib import Path

from tools.validate_crud_coverage_index import (
    ValidationError,
    _validate_identity_artifact,
    validate_index,
    validate_paths,
)


ROOT = Path(__file__).resolve().parents[1]
INDEX_PATH = ROOT / "docs/performance/crud-coverage-index-v1.json"
CATALOG_PATH = ROOT / "docs/performance/results/perf-corpus-manifest-v2.json"
SELECTOR_PATH = ROOT / "tools/perf-baseline/src/lib.rs"
CHECKLIST_PATH = ROOT / "docs/CRUD_Scenario_Checklist.md"
EXPECTED_CATALOG_SHA256 = (
    "679d39548de0150b9fdfcdd2628ca3adc1bbd3c041a21d7a95ab70095c2d0ba9"
)


def _canonical(value: object) -> bytes:
    return json.dumps(
        value,
        ensure_ascii=False,
        allow_nan=False,
        sort_keys=True,
        separators=(",", ":"),
    ).encode("utf-8")


def _refresh_catalog_hashes(catalog: dict) -> None:
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
    catalog["content_set_sha256"] = hashlib.sha256(_canonical(content_set)).hexdigest()
    without_hash = dict(catalog)
    without_hash.pop("catalog_sha256")
    catalog["catalog_sha256"] = hashlib.sha256(_canonical(without_hash)).hexdigest()


class CrudCoverageIndexTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.index = json.loads(INDEX_PATH.read_text(encoding="utf-8"))
        cls.catalog = json.loads(CATALOG_PATH.read_text(encoding="utf-8"))
        cls.identity = json.loads(
            (ROOT / "docs/performance/results/perf-regression-default-manifest-v1.json").read_text(
                encoding="utf-8"
            )
        )
        cls.selector_source = SELECTOR_PATH.read_text(encoding="utf-8")
        cls.checklist_source = CHECKLIST_PATH.read_text(encoding="utf-8")

    def validate(
        self,
        index: dict,
        catalog: dict | None = None,
        *,
        checked_catalog: dict | None = None,
        report: dict | None = None,
        selector_source: str | None = None,
    ) -> tuple[int, int]:
        return validate_index(
            index,
            self.catalog if catalog is None else catalog,
            self.selector_source if selector_source is None else selector_source,
            self.checklist_source,
            repo_root=ROOT,
            checked_catalog=checked_catalog,
            report=report,
        )

    def measured_report(self, catalog: dict | None = None) -> dict:
        catalog = self.catalog if catalog is None else catalog
        corpora = {corpus["id"]: corpus for corpus in catalog["corpora"]}
        results = []
        for category in self.index["categories"]:
            if category["status"] != "measured":
                continue
            for scenario in category["scenarios"]:
                for identifier in scenario["corpus"]["ids"]:
                    results.append(
                        {
                            "case": scenario["selector"],
                            "corpus": corpora[identifier]["legacy_v1"],
                            "elapsed_ns": {
                                "unit": "ns",
                                "samples": [100] * 15,
                                "sample_order": list(range(15)),
                                "min": 100,
                                "p50": 100,
                                "p95": 100,
                                "p99": 100,
                                "max": 100,
                                "mean": 100.0,
                                "standard_deviation": 0.0,
                                "confidence_interval_95": {
                                    "method": "two-sided Student's t interval for the mean",
                                    "lower": 100.0,
                                    "upper": 100.0,
                                },
                            },
                        }
                    )
        return {
            "schema_version": 1,
            "corpus_catalog": {
                "manifest_version": catalog["manifest_version"],
                "catalog_id": catalog["catalog_id"],
                "catalog_sha256": catalog["catalog_sha256"],
                "content_set_sha256": catalog["content_set_sha256"],
            },
            "results": results,
        }

    def test_checked_index_has_all_fifteen_categories_and_real_selectors(self) -> None:
        self.assertEqual(self.validate(self.index), (15, 30))

    def test_checked_catalog_hash_matches_regenerated_catalog(self) -> None:
        checked_catalog = self.index["checked_catalog"]
        self.assertEqual(checked_catalog["catalog_sha256"], EXPECTED_CATALOG_SHA256)
        self.assertEqual(checked_catalog["catalog_sha256"], self.catalog["catalog_sha256"])

    def test_missing_category_is_rejected(self) -> None:
        index = copy.deepcopy(self.index)
        index["categories"].pop()
        with self.assertRaises(ValidationError):
            self.validate(index)

    def test_unknown_selector_is_rejected(self) -> None:
        index = copy.deepcopy(self.index)
        index["categories"][0]["scenarios"][0]["selector"] = "invented_case"
        with self.assertRaises(ValidationError):
            self.validate(index)

    def test_duplicate_selector_is_rejected(self) -> None:
        index = copy.deepcopy(self.index)
        index["categories"][1]["scenarios"][0]["selector"] = (
            index["categories"][0]["scenarios"][0]["selector"]
        )
        with self.assertRaises(ValidationError):
            self.validate(index)

    def test_checked_corpus_case_must_match_selector(self) -> None:
        index = copy.deepcopy(self.index)
        index["categories"][0]["scenarios"][0]["corpus"]["case"] = "xlsx_list_sheets"
        with self.assertRaises(ValidationError):
            self.validate(index)

    def test_checked_corpus_id_must_match_every_catalog_binding(self) -> None:
        index = copy.deepcopy(self.index)
        ids = index["categories"][0]["scenarios"][0]["corpus"]["ids"]
        ids.pop()
        with self.assertRaises(ValidationError):
            self.validate(index)

    def test_generated_per_run_identity_is_required(self) -> None:
        index = copy.deepcopy(self.index)
        identity = index["categories"][2]["scenarios"][0]["corpus"]["identity"]
        index["categories"][2]["scenarios"][0]["corpus"]["identity"] = identity.replace(
            "schema-2", "schema-one"
        )
        with self.assertRaises(ValidationError):
            self.validate(index)

    def test_status_requires_evidence_and_selector_shape(self) -> None:
        index = copy.deepcopy(self.index)
        del index["categories"][0]["scenarios"][0]["evidence"]
        with self.assertRaises(ValidationError):
            self.validate(index)

    def test_category_status_is_derived_from_scenario_statuses(self) -> None:
        index = copy.deepcopy(self.index)
        index["categories"][0]["scenarios"][0]["status"] = "correctness-only"
        with self.assertRaises(ValidationError):
            self.validate(index)

    def test_measured_status_requires_retained_timing_artifact(self) -> None:
        index = copy.deepcopy(self.index)
        index["categories"][0]["measurement"] = {
            "kind": "retained-timing",
            "artifact": "docs/performance/results/perf-regression-default-manifest-v1.json",
        }
        with self.assertRaises(ValidationError):
            self.validate(index)

    def test_measured_status_requires_retained_default_baseline(self) -> None:
        index = copy.deepcopy(self.index)
        index["categories"][0]["measurement"]["identity_artifact"] = (
            "tools/perf-baseline/README.md"
        )
        with self.assertRaises(ValidationError):
            self.validate(index)

    def test_measured_status_rejects_generated_per_run_corpus(self) -> None:
        index = copy.deepcopy(self.index)
        index["categories"][0]["scenarios"][0]["corpus"] = {
            "kind": "generated-per-run",
            "shapes": ["tiny"],
            "identity": "harness-generated schema-2 corpus IDs use format:sha256:<archive_sha256>",
        }
        with self.assertRaises(ValidationError):
            self.validate(index)

    def test_correctness_only_status_explains_absent_timing_claim(self) -> None:
        index = copy.deepcopy(self.index)
        index["categories"][2]["measurement"]["reason"] = (
            "The selectors are retained as correctness evidence."
        )
        with self.assertRaises(ValidationError):
            self.validate(index)

    def test_generated_shape_and_identity_contracts_are_exact(self) -> None:
        index = copy.deepcopy(self.index)
        corpus = index["categories"][2]["scenarios"][0]["corpus"]
        corpus["shapes"] = []
        with self.assertRaises(ValidationError):
            self.validate(index)
        index = copy.deepcopy(self.index)
        corpus = index["categories"][2]["scenarios"][0]["corpus"]
        corpus["shapes"] = ["bogus"]
        with self.assertRaises(ValidationError):
            self.validate(index)
        index = copy.deepcopy(self.index)
        corpus = index["categories"][2]["scenarios"][0]["corpus"]
        corpus["identity"] = "schema-2 sha256"
        with self.assertRaises(ValidationError):
            self.validate(index)

    def test_index_and_category_claim_representative_coverage(self) -> None:
        index = copy.deepcopy(self.index)
        index["coverage_claim"] = "complete"
        with self.assertRaises(ValidationError):
            self.validate(index)
        index = copy.deepcopy(self.index)
        index["categories"][0]["coverage_scope"] = "complete"
        with self.assertRaises(ValidationError):
            self.validate(index)

    def test_list_valued_checklist_field_is_rejected_as_validation_error(self) -> None:
        index = copy.deepcopy(self.index)
        index["categories"][0]["checklist_refs"] = "Open valid input"
        with self.assertRaises(ValidationError):
            self.validate(index)

    def test_missing_nested_corpus_field_is_rejected_as_validation_error(self) -> None:
        index = copy.deepcopy(self.index)
        del index["categories"][2]["scenarios"][0]["corpus"]["identity"]
        with self.assertRaises(ValidationError):
            self.validate(index)

    def test_scope_cannot_include_iwork(self) -> None:
        index = copy.deepcopy(self.index)
        index["scope"]["excluded_formats"] = []
        with self.assertRaises(ValidationError):
            self.validate(index)
        index["scope"]["excluded_formats"] = ["iWork"]
        index["scope"]["included_formats"] = "iWork"
        with self.assertRaises(ValidationError):
            self.validate(index)

    def test_checklist_refs_must_be_nonempty(self) -> None:
        index = copy.deepcopy(self.index)
        index["categories"][0]["checklist_refs"] = []
        with self.assertRaises(ValidationError):
            self.validate(index)

    def test_navigation_paths_are_approved_documentation_only(self) -> None:
        index = copy.deepcopy(self.index)
        index["categories"][0]["scenarios"][0]["evidence"] = ["Cargo.toml"]
        with self.assertRaises(ValidationError):
            self.validate(index)

    def test_unrelated_iwork_selector_in_registry_is_allowed(self) -> None:
        marker = "\n}\n\nimpl Case"
        selector_source = self.selector_source.replace(
            marker,
            "\n    FutureIworkCase,\n}" + "\n\nimpl Case",
            1,
        )
        selector_source = selector_source.replace(
            'Self::XlsxFullCellScan => "xlsx_full_cell_scan",',
            'Self::XlsxFullCellScan => "xlsx_full_cell_scan",\n            Self::FutureIworkCase => "future_iwork_case",',
            1,
        )
        self.assertEqual(self.validate(self.index, selector_source=selector_source), (15, 30))

    def test_referenced_iwork_selector_is_rejected(self) -> None:
        marker = "\n}\n\nimpl Case"
        selector_source = self.selector_source.replace(
            marker,
            "\n    FutureIworkCase,\n}" + "\n\nimpl Case",
            1,
        )
        selector_source = selector_source.replace(
            'Self::XlsxFullCellScan => "xlsx_full_cell_scan",',
            'Self::XlsxFullCellScan => "xlsx_full_cell_scan",\n            Self::FutureIworkCase => "future_iwork_case",',
            1,
        )
        index = copy.deepcopy(self.index)
        index["categories"][0]["scenarios"][0]["selector"] = "future_iwork_case"
        with self.assertRaises(ValidationError):
            self.validate(index, selector_source=selector_source)

    def test_measured_binding_role_must_be_timed_after_rehash(self) -> None:
        catalog = copy.deepcopy(self.catalog)
        binding = next(
            item
            for item in catalog["case_bindings"]
            if item["case"] == "xlsx_full_cell_scan"
        )
        binding["role"] = "guard"
        _refresh_catalog_hashes(catalog)
        with self.assertRaises(ValidationError):
            self.validate(self.index, catalog=catalog)

    def test_identity_artifact_digest_and_shape_are_strict(self) -> None:
        identity = copy.deepcopy(self.identity)
        identity["result_keys_sha256"] = "0" * 64
        with self.assertRaises(ValidationError):
            _validate_identity_artifact(identity, "identity")
        identity = copy.deepcopy(self.identity)
        identity["default_cases"].pop()
        with self.assertRaises(ValidationError):
            _validate_identity_artifact(identity, "identity")

    def test_duplicate_json_keys_are_rejected_by_validate_paths(self) -> None:
        original = INDEX_PATH.read_text(encoding="utf-8")
        duplicate = original.replace(
            '"schema_version": 1,', '"schema_version": 1,\n  "schema_version": 1,', 1
        )
        with tempfile.TemporaryDirectory() as directory:
            index_path = Path(directory) / "index.json"
            index_path.write_text(duplicate, encoding="utf-8")
            with self.assertRaises(ValidationError):
                validate_paths(
                    index_path,
                    CATALOG_PATH,
                    SELECTOR_PATH,
                    CHECKLIST_PATH,
                    repo_root=ROOT,
                )

    def test_catalog_semantic_tamper_is_rejected_even_if_index_is_unchanged(self) -> None:
        catalog = copy.deepcopy(self.catalog)
        corpus = catalog["corpora"][0]
        original = corpus["bytes"]["archive_sha256"]
        corpus["bytes"]["archive_sha256"] = "0" * 64 if original != "0" * 64 else "1" * 64
        with self.assertRaises(ValidationError):
            validate_index(
                self.index,
                catalog,
                self.selector_source,
                self.checklist_source,
                repo_root=ROOT,
            )

    def test_representative_report_binds_every_measured_row(self) -> None:
        report = self.measured_report()
        self.assertEqual(self.validate(self.index, report=report), (15, 30))

    def test_validate_paths_binds_report_to_a_separate_run_catalog(self) -> None:
        report = self.measured_report()
        with tempfile.TemporaryDirectory() as directory:
            directory_path = Path(directory)
            report_path = directory_path / "report.json"
            catalog_path = directory_path / "catalog.json"
            report_path.write_text(json.dumps(report), encoding="utf-8")
            catalog_path.write_text(json.dumps(self.catalog), encoding="utf-8")
            self.assertEqual(
                validate_paths(
                    INDEX_PATH,
                    catalog_path,
                    SELECTOR_PATH,
                    CHECKLIST_PATH,
                    repo_root=ROOT,
                    report_path=report_path,
                ),
                (15, 30),
            )

    def test_report_missing_measured_row_is_rejected(self) -> None:
        report = self.measured_report()
        report["results"].pop()
        with self.assertRaises(ValidationError):
            self.validate(self.index, report=report)

    def test_report_malformed_case_and_elapsed_statistics_are_rejected(self) -> None:
        report = self.measured_report()
        report["results"][0]["case"] = [report["results"][0]["case"]]
        with self.assertRaises(ValidationError):
            self.validate(self.index, report=report)
        report = self.measured_report()
        report["results"][0]["elapsed_ns"]["sample_order"][0] = 1
        with self.assertRaises(ValidationError):
            self.validate(self.index, report=report)

    def test_report_rejects_negative_u64_values_even_with_matching_stats(self) -> None:
        report = self.measured_report()
        elapsed = report["results"][0]["elapsed_ns"]
        elapsed["samples"] = [-1] * 15
        elapsed["min"] = -1
        elapsed["p50"] = -1
        elapsed["p95"] = -1
        elapsed["p99"] = -1
        elapsed["max"] = -1
        elapsed["mean"] = -1.0
        elapsed["standard_deviation"] = 0.0
        elapsed["confidence_interval_95"]["lower"] = -1.0
        elapsed["confidence_interval_95"]["upper"] = -1.0
        with self.assertRaises(ValidationError):
            self.validate(self.index, report=report)

        for summary in ("min", "p50", "p95", "p99", "max"):
            report = self.measured_report()
            report["results"][0]["elapsed_ns"][summary] = -1
            with self.assertRaises(ValidationError):
                self.validate(self.index, report=report)

    def test_report_recomputes_statistics_and_tie_order(self) -> None:
        mutations = (
            ("mean", 101.0),
            ("standard_deviation", 1.0),
        )
        for key, value in mutations:
            report = self.measured_report()
            report["results"][0]["elapsed_ns"][key] = value
            with self.assertRaises(ValidationError):
                self.validate(self.index, report=report)

        report = self.measured_report()
        report["results"][0]["elapsed_ns"]["confidence_interval_95"]["lower"] = 99.0
        with self.assertRaises(ValidationError):
            self.validate(self.index, report=report)

        report = self.measured_report()
        report["results"][0]["elapsed_ns"]["confidence_interval_95"]["method"] = "fixture"
        with self.assertRaises(ValidationError):
            self.validate(self.index, report=report)

        report = self.measured_report()
        report["results"][0]["elapsed_ns"]["sample_order"][:2] = [1, 0]
        with self.assertRaises(ValidationError):
            self.validate(self.index, report=report)

    def test_report_catalog_reference_must_match_run_catalog(self) -> None:
        report = self.measured_report()
        report["corpus_catalog"]["content_set_sha256"] = "0" * 64
        with self.assertRaises(ValidationError):
            self.validate(self.index, report=report)


if __name__ == "__main__":
    unittest.main()
