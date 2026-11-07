"""Integrity of the native Office corpus; this is not semantic API coverage.

Run with ``python3 -m unittest tools.test_native_office_corpus``. The native
authoring inventory remains historical evidence; Rust regression tests must
separately establish the behavior of the current library.
"""

from collections import Counter
import copy
import hashlib
import json
from pathlib import Path
import unittest
import zipfile

from tools.native_office_coverage import EVIDENCE_REVIEWS, evidence_reviews, report
from tools.non_iwork_gate import FACADE_FORMAT_TEST_FEATURES


ROOT = Path(__file__).resolve().parents[1]
CORPUS = ROOT / "test-data/microsoft-office"


def document(name: str) -> dict:
    return json.loads((CORPUS / name).read_text(encoding="utf-8"))


class NativeOfficeCorpus(unittest.TestCase):
    def test_manifest_and_assets_match_saved_native_bytes(self) -> None:
        manifest = document("manifest.json")
        fixtures = manifest["files"]
        assets = manifest["import_asset_manifest"]["files"]
        self.assertEqual(len(fixtures), 139)
        self.assertEqual(len(assets), 4)
        self.assertEqual(len({item["path"] for item in fixtures}), len(fixtures))
        self.assertEqual(len({item["path"] for item in assets}), len(assets))
        for base, entries in ((CORPUS, fixtures), (ROOT, assets)):
            for item in entries:
                with self.subTest(path=item["path"]):
                    path = base / item["path"]
                    self.assertTrue(path.resolve().is_relative_to(CORPUS.resolve()))
                    data = path.read_bytes()
                    self.assertEqual(len(data), item["bytes"])
                    self.assertEqual(hashlib.sha256(data).hexdigest(), item["sha256"])

    def test_every_fixture_has_inspection_and_native_evidence(self) -> None:
        fixtures = {
            "test-data/microsoft-office/" + item["path"]: item
            for item in document("manifest.json")["files"]
        }
        coverage = document("coverage.json")
        for entries in (
            document("inspection.json")["fixtures"],
            coverage["native_fixture_evidence"],
            coverage["current_native_fixtures"],
        ):
            self.assertEqual(len(entries), len(fixtures))
            self.assertEqual({item["path"] for item in entries}, set(fixtures))
            for item in entries:
                with self.subTest(path=item["path"]):
                    expected = fixtures[item["path"]]
                    self.assertEqual(item["sha256"], expected["sha256"])
                    if "size_bytes" in item:
                        self.assertEqual(item["size_bytes"], expected["bytes"])

    def test_native_containers_are_intact(self) -> None:
        counts = Counter()
        for item in document("manifest.json")["files"]:
            with self.subTest(path=item["path"]):
                path = CORPUS / item["path"]
                prefix = path.read_bytes()[:8]
                if prefix.startswith(b"PK\x03\x04"):
                    with zipfile.ZipFile(path) as archive:
                        self.assertIsNone(archive.testzip())
                        self.assertIn("[Content_Types].xml", archive.namelist())
                    counts["opc"] += 1
                elif prefix == bytes.fromhex("d0cf11e0a1b11ae1"):
                    counts["cfb"] += 1
                else:
                    self.assertTrue(prefix.startswith(b"{\\rtf1"))
                    counts["rtf"] += 1
        self.assertEqual(counts, {"opc": 71, "cfb": 53, "rtf": 15})

    def test_feature_inventory_and_pending_dispositions_are_complete(self) -> None:
        coverage = document("coverage.json")
        inventory = coverage["feature_inventory"]
        by_id = {item["id"]: item for item in inventory}
        self.assertEqual(len(inventory), 494)
        self.assertEqual(len(by_id), len(inventory))
        self.assertEqual(
            {item["format"] for item in inventory},
            {"doc", "docx", "rtf", "xls", "xlsx", "xlsb", "ppt", "pptx"},
        )
        review = document("final-coverage-review.json")
        snapshot_ids = []
        for status, snapshot in review["status_snapshot"].items():
            expected = {item["id"] for item in inventory if item["fixture_status"] == status}
            self.assertEqual(set(snapshot["ids"]), expected, status)
            self.assertEqual(len(snapshot["ids"]), snapshot["count"])
            snapshot_ids.extend(snapshot["ids"])
        self.assertEqual(Counter(snapshot_ids), Counter(by_id.keys()))
        pending = {item["id"] for item in inventory if item["fixture_status"] == "pending"}
        dispositions = coverage["residual_matrix_pending_categories"]
        assigned = [row for group in dispositions["categories"].values() for row in group["row_ids"]]
        self.assertEqual(Counter(assigned), Counter(pending))
        self.assertEqual(len(assigned), dispositions["categorized_pending_count"])
        self.assertEqual(len(pending), dispositions["pending_count"])
        self.assertEqual(Counter(item["id"] for item in review["pending_rows"]), Counter(pending))
        for item in review["pending_rows"]:
            self.assertTrue(item["rationale"].strip(), item["id"])

    def test_all_corpus_files_are_inventoried_or_json_evidence(self) -> None:
        manifest = document("manifest.json")
        payloads = {CORPUS / item["path"] for item in manifest["files"]}
        payloads.update(ROOT / item["path"] for item in manifest["import_asset_manifest"]["files"])
        for path in CORPUS.rglob("*"):
            if path.is_file() and path not in payloads:
                with self.subTest(path=str(path.relative_to(CORPUS))):
                    self.assertEqual(path.suffix, ".json", "add new payloads to the manifest")
                    self.assertIsInstance(json.loads(path.read_text(encoding="utf-8")), dict)

    def test_test_index_accounts_for_uncovered_fixtures_and_feature_rows(self) -> None:
        coverage = report()
        self.assertEqual(len(coverage["fixtures"]), 139)
        self.assertEqual(len(coverage["feature_rows"]), 494)
        for row in coverage["feature_rows"]:
            self.assertEqual(row["automated_coverage"], "partial" if row["tests"] else "uncovered")

    def test_facade_gate_enables_every_indexed_native_test(self) -> None:
        tests = [test for test in report()["tests"] if test["crate"] == "litchi"]
        self.assertTrue(tests)
        for test in tests:
            self.assertTrue(set(test["required_features"]) <= set(FACADE_FORMAT_TEST_FEATURES), test["name"])

    def test_gap_report_does_not_promote_native_evidence_or_pending_exclusions(self) -> None:
        coverage = report()
        counts = Counter(row["verification_gap"] for row in coverage["feature_rows"])
        self.assertEqual(sum(counts.values()), 494)
        self.assertEqual(dict(counts), coverage["summary"]["verification_gaps"])
        for row in coverage["feature_rows"]:
            with self.subTest(row=row["id"]):
                if row["tests"]:
                    self.assertEqual(row["verification_gap"], "partial_assertions_require_scope_review")
                elif "evidence_review" in row:
                    self.assertEqual(row["verification_gap"], "reviewed_native_fixture_gap_requires_followup")
                elif row["native_fixture_status"] == "verified_native_fixture":
                    self.assertEqual(row["verification_gap"], "native_evidence_needs_semantic_tests")
                elif row["native_fixture_status"] == "pending":
                    self.assertTrue(row["historical_evidence_refs"])
                    self.assertNotIn(row["verification_gap"], {"not_applicable", "complete"})

    def test_evidence_reviews_never_count_absence_as_semantic_coverage(self) -> None:
        coverage = report()
        reviewed = [row for row in coverage["feature_rows"] if "evidence_review" in row]
        self.assertEqual({row["id"] for row in reviewed}, {"xlsx-240", "xlsx-245", "xlsx-266", "ppt-363", "docx-093", "docx-097", "ppt-341", "ppt-379", "xls-176"})
        self.assertEqual(coverage["summary"]["reviewed_native_fixture_gaps"], len(reviewed))
        historical = {row["id"]: row["fixture_status"] for row in document("coverage.json")["feature_inventory"]}
        for row in reviewed:
            with self.subTest(row=row["id"]):
                self.assertEqual(row["automated_coverage"], "uncovered")
                self.assertEqual(row["tests"], [])
                self.assertEqual(row["native_fixture_status"], historical[row["id"]])
                self.assertTrue(row["evidence_review"]["checks"])

    def test_evidence_review_validation_rejects_unknown_or_unsubstantiated_claims(self) -> None:
        original = json.loads(EVIDENCE_REVIEWS.read_text())
        rows = {item["id"]: [] for item in document("coverage.json")["feature_inventory"]}
        fixtures = {item["path"]: [] for item in document("manifest.json")["files"]}
        for case in ("duplicate", "unknown_row", "unknown_fixture", "empty_reason", "missing_check", "missing_source", "completion_claim"):
            with self.subTest(case=case):
                candidate = copy.deepcopy(original)
                review = candidate["reviews"][0]
                if case == "duplicate":
                    candidate["reviews"].append(copy.deepcopy(review))
                elif case == "unknown_row":
                    review["feature_row"] = "xlsx-999"
                elif case == "unknown_fixture":
                    review["fixtures"] = ["excel/not-a-native-fixture.xlsx"]
                elif case == "empty_reason":
                    review["reason"] = " "
                elif case == "missing_check":
                    review["checks"][0]["name"] = "missing_evidence_check"
                elif case == "missing_source":
                    review["evidence_refs"] = ["missing-evidence.json"]
                else:
                    review["classification"] = "complete"
                with self.assertRaises(ValueError):
                    evidence_reviews(candidate, rows, fixtures)

    def test_fixture_gap_cannot_gain_semantic_credit_without_reconciliation(self) -> None:
        import tempfile
        from unittest.mock import patch
        from tools.native_office_coverage import INDEX

        index = json.loads(INDEX.read_text())
        index["tests"][0]["feature_rows"].append("xlsx-240")
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "index.json"
            path.write_text(json.dumps(index))
            with patch("tools.native_office_coverage.INDEX", path):
                with self.assertRaisesRegex(ValueError, "fixture-gap review conflicts"):
                    report()


if __name__ == "__main__":
    unittest.main()
