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

from tools.native_office_coverage import (
    EVIDENCE_REVIEWS, evidence_reviews, matrix_rows, reconcile_matrices, report,
)
from tools.non_iwork_gate import FACADE_FORMAT_TEST_FEATURES


ROOT = Path(__file__).resolve().parents[1]
CORPUS = ROOT / "test-data/microsoft-office"


def document(name: str) -> dict:
    return json.loads((CORPUS / name).read_text(encoding="utf-8"))


class NativeOfficeCorpus(unittest.TestCase):
    def test_index_accepts_registered_async_tests_but_rejects_plain_helpers(self) -> None:
        import tempfile
        from unittest.mock import patch
        from tools.native_office_coverage import test_identity

        check = {"crate": "litchi", "target": "coordinates", "name": "native_cells"}
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source = root / "crates/litchi/tests/coordinates.rs"
            source.parent.mkdir(parents=True)
            with patch("tools.native_office_coverage.ROOT", root):
                source.write_text("#[tokio::test]\nasync fn native_cells() {}\n")
                self.assertEqual(test_identity(check), "litchi::coordinates::native_cells")
                for body in ("async fn native_cells() {}", "#[test]\nfn different_test() {}"):
                    source.write_text(body)
                    with self.assertRaisesRegex(ValueError, "missing test"):
                        test_identity(check)

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

    def test_current_checklists_are_reconciled_without_rewriting_historical_evidence(self) -> None:
        coverage = report()
        current = coverage["current_matrix"]
        self.assertEqual(len(current["rows"]), 495)
        self.assertEqual(len(current["renamed_rows"]), 2)
        self.assertEqual(len(current["consolidated_rows"]), 1)
        self.assertEqual(len(current["status_changes"]), 13)
        additions = [row for row in current["rows"] if row["historical_feature_row"] is None]
        self.assertEqual({row["id"] for row in additions}, {
            "docx-current-direct-properties", "xlsb-current-fresh-creation",
        })
        for row in additions:
            self.assertTrue(row["tests"])
            self.assertEqual(row["automated_coverage"], "partial")
            self.assertNotIn("native_fixture_status", row)
        historical = {row["id"]: row for row in coverage["feature_rows"]}
        self.assertEqual(len(historical), 494)
        self.assertEqual(historical["xlsx-264"]["verification_gap"], "reviewed_native_fixture_gap_requires_followup")
        self.assertNotIn("xlsx-264", {row["id"] for row in current["rows"]})
        self.assertEqual(historical["rtf-455"]["feature"], "Immutable snapshot and borrowed story/resource views")
        for row in current["rows"]:
            if row["historical_feature_row"] is not None:
                old = historical[row["historical_feature_row"]]
                self.assertEqual(row["tests"], old["tests"])
                self.assertEqual(row["verification_gap"], old["verification_gap"])

    def test_additional_feature_test_bindings_require_known_ids_and_real_tests(self) -> None:
        import tempfile
        from unittest.mock import patch
        from tools.native_office_coverage import INDEX

        original = json.loads(INDEX.read_text())
        for case in ("unknown", "historical_id", "empty", "missing_test"):
            with self.subTest(case=case), tempfile.TemporaryDirectory() as directory:
                index = copy.deepcopy(original)
                test = next(test for test in index["tests"] if test.get("current_features"))
                if case == "missing_test":
                    test["name"] = "nonexistent_current_feature_test"
                else:
                    test["current_features"] = {
                        "unknown": ["docx-current-unknown"],
                        "historical_id": ["docx-069"],
                        "empty": [],
                    }[case]
                path = Path(directory) / "index.json"
                path.write_text(json.dumps(index))
                with patch("tools.native_office_coverage.INDEX", path):
                    with self.assertRaises(ValueError):
                        report()

    def test_facade_gate_enables_every_indexed_native_test(self) -> None:
        tests = [test for test in report()["tests"] if test["crate"] == "litchi"]
        self.assertTrue(tests)
        for test in tests:
            self.assertTrue(set(test["required_features"]) <= set(FACADE_FORMAT_TEST_FEATURES), test["name"])

    def test_unsupported_ingress_has_boundary_checks_without_positive_feature_credit(self) -> None:
        coverage = report()
        self.assertEqual(coverage["summary"]["fixtures_with_any_checks"], 139)
        for fixture in coverage["fixtures"]:
            self.assertTrue(fixture["tests"] or fixture["evidence_checks"], fixture["path"])
        encrypted = next(fixture for fixture in coverage["fixtures"] if fixture["path"] == "excel/excel-encrypted.xlsb")
        self.assertEqual(encrypted["tests"], [])
        self.assertTrue(any("unsupported_container_boundary" in check for check in encrypted["evidence_checks"]))
        rows = {row["id"]: row for row in coverage["feature_rows"]}
        for identity in ("doc-065", "xlsx-268", "xlsb-329"):
            self.assertEqual(rows[identity]["verification_gap"], "reviewed_explicit_scope_boundary")
            self.assertEqual(rows[identity]["tests"], [])

    def test_unsupported_current_scope_cannot_gain_credit_by_deleting_its_review(self) -> None:
        import tempfile
        from unittest.mock import patch
        from tools.native_office_coverage import INDEX

        index = json.loads(INDEX.read_text())
        index["tests"][0]["feature_rows"].append("doc-065")
        evidence = json.loads(EVIDENCE_REVIEWS.read_text())
        evidence["reviews"] = [row for row in evidence["reviews"] if row["feature_row"] != "doc-065"]
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            (root / "index.json").write_text(json.dumps(index))
            (root / "reviews.json").write_text(json.dumps(evidence))
            with patch("tools.native_office_coverage.INDEX", root / "index.json"), patch(
                "tools.native_office_coverage.EVIDENCE_REVIEWS", root / "reviews.json"
            ):
                with self.assertRaisesRegex(ValueError, "unsupported matrix feature"):
                    report()

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
                    expected = (
                        "reviewed_explicit_scope_boundary"
                        if row["evidence_review"]["classification"] == "explicit_scope_boundary"
                        else "reviewed_native_fixture_gap_requires_followup"
                    )
                    self.assertEqual(row["verification_gap"], expected)
                elif row["native_fixture_status"] == "verified_native_fixture":
                    self.assertEqual(row["verification_gap"], "native_evidence_needs_semantic_tests")
                elif row["native_fixture_status"] == "pending":
                    self.assertTrue(row["historical_evidence_refs"])
                    self.assertNotIn(row["verification_gap"], {"not_applicable", "complete"})

    def test_evidence_reviews_never_count_absence_as_semantic_coverage(self) -> None:
        coverage = report()
        reviewed = [row for row in coverage["feature_rows"] if "evidence_review" in row]
        self.assertEqual({row["id"] for row in reviewed}, {
            "xlsx-240", "xlsx-245", "xlsx-266", "ppt-363", "docx-093", "docx-097",
            "ppt-341", "ppt-379", "xls-176", "ppt-355", "doc-014", "doc-016",
            "doc-059", "xls-173", "xls-178", "ppt-377", "ppt-383", "ppt-384",
            "pptx-432", "pptx-435", "pptx-449", "ppt-343", "ppt-364",
            "doc-068", "docx-133", "pptx-451", "xls-193", "xlsx-270", "xlsb-331",
            "doc-061", "docx-123", "docx-125",
            "docx-130", "xls-194", "xlsx-269", "xlsb-334", "ppt-375", "ppt-393",
            "pptx-452", "rtf-493",
            "doc-022", "xls-196", "xlsx-273", "xlsb-339", "docx-132",
            "xls-195", "xlsx-271", "xlsb-338",
            "doc-064", "docx-127", "xlsx-272", "xlsb-337", "ppt-392", "rtf-494",
            "xlsb-336", "docx-126",
            "doc-067", "xls-189", "xls-190", "xls-191", "xlsx-267", "xlsb-332",
            "ppt-369", "ppt-390",
            "doc-002", "doc-003", "doc-004", "doc-005", "doc-006", "doc-010", "doc-015",
            "pptx-438",
            "rtf-486", "rtf-489", "rtf-491",
            "xlsx-200", "xlsx-264", "xlsb-286", "xlsb-287",
            "docx-086",
            "doc-052",
            "ppt-389", "ppt-391",
            "doc-065", "xlsx-268", "xlsb-329",
        })
        boundaries = [row for row in reviewed if row["evidence_review"]["classification"] == "explicit_scope_boundary"]
        self.assertEqual({row["id"] for row in boundaries}, {
            "doc-068", "docx-133", "pptx-451", "xls-193", "xlsx-270", "xlsb-331",
            "doc-061", "docx-123", "docx-125",
            "docx-130", "xls-194", "xlsx-269", "xlsb-334", "ppt-375", "ppt-393",
            "pptx-452", "rtf-493",
            "doc-022", "xls-196", "xlsx-273", "xlsb-339", "docx-132",
            "xls-195", "xlsx-271", "xlsb-338",
            "doc-064", "docx-127", "xlsx-272", "xlsb-337", "ppt-392", "rtf-494",
            "xlsb-336", "docx-126",
            "doc-067", "xls-189", "xls-190", "xls-191", "xlsx-267", "xlsb-332",
            "ppt-369", "ppt-390",
            "ppt-389", "ppt-391",
            "doc-065", "xlsx-268", "xlsb-329",
        })
        self.assertEqual(coverage["summary"]["reviewed_native_fixture_gaps"], len(reviewed) - len(boundaries))
        self.assertEqual(coverage["summary"]["reviewed_explicit_scope_boundaries"], len(boundaries))
        historical = {row["id"]: row["fixture_status"] for row in document("coverage.json")["feature_inventory"]}
        for row in reviewed:
            with self.subTest(row=row["id"]):
                self.assertEqual(row["automated_coverage"], "uncovered")
                self.assertEqual(row["tests"], [])
                self.assertEqual(row["native_fixture_status"], historical[row["id"]])
                self.assertTrue(row["evidence_review"]["checks"])

    def test_word_refresh_is_not_excluded_by_worksheet_evaluator_refusals(self) -> None:
        rows = {row["id"]: row for row in report()["feature_rows"]}
        for name in ("doc-038", "doc-066", "docx-124"):
            with self.subTest(row=name):
                self.assertNotIn("evidence_review", rows[name])
                self.assertEqual(rows[name]["automated_coverage"], "partial")
                self.assertTrue(any("native_caption_caches_refresh_in_doc_docx_and_templates_without_layout" in test for test in rows[name]["tests"]))

    def test_scope_boundary_requires_both_sides_of_the_contract(self) -> None:
        original = json.loads(EVIDENCE_REVIEWS.read_text())
        rows = {item["id"]: [] for item in document("coverage.json")["feature_inventory"]}
        fixtures = {item["path"]: [] for item in document("manifest.json")["files"]}
        for field in ("supported_scope", "unsupported_scope"):
            for value in (None, "", " "):
                with self.subTest(field=field, value=value):
                    candidate = copy.deepcopy(original)
                    boundary = next(review for review in candidate["reviews"] if review["classification"] == "explicit_scope_boundary")
                    if value is None:
                        del boundary[field]
                    else:
                        boundary[field] = value
                    with self.assertRaisesRegex(ValueError, field):
                        evidence_reviews(candidate, rows, fixtures)

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

    def test_evidence_review_cannot_gain_semantic_credit_without_reconciliation(self) -> None:
        import tempfile
        from unittest.mock import patch
        from tools.native_office_coverage import INDEX

        for row in ("xlsx-240", "doc-068"):
            with self.subTest(row=row):
                index = json.loads(INDEX.read_text())
                index["tests"][0]["feature_rows"].append(row)
                with tempfile.TemporaryDirectory() as directory:
                    path = Path(directory) / "index.json"
                    path.write_text(json.dumps(index))
                    with patch("tools.native_office_coverage.INDEX", path):
                        with self.assertRaisesRegex(ValueError, "evidence review conflicts"):
                            report()


class MatrixReconciliation(unittest.TestCase):
    def setUp(self) -> None:
        import tempfile
        from unittest.mock import patch

        temporary = tempfile.TemporaryDirectory()
        self.addCleanup(temporary.cleanup)
        self.root = Path(temporary.name)
        root_patch = patch("tools.native_office_coverage.ROOT", self.root)
        root_patch.start()
        self.addCleanup(root_patch.stop)
        self.matrix = self.root / "matrix.md"
        self.matrix.write_text(
            "| Feature | Status | Read | Write |\n| --- | --- | --- | --- |\n"
            "| Base | 🟡 | ✅ | 🟡 |\n| Renamed | ✅ | ✅ | ✅ |\n"
            "| Added | ✅ | ✅ | ✅ |\n"
        )
        self.inventory = [
            {"id": f"docx-{number}", "format": "docx", "source_matrix": "matrix.md",
             "feature": feature, "matrix_status": "✅", "read": "✅", "write": "✅"}
            for number, feature in enumerate(("Base", "Old name", "Consolidated"), 1)
        ]
        self.document = {
            "schema_version": 1,
            "renamed_rows": [{"feature_row": "docx-2", "current_feature": "Renamed", "reason": "API naming."}],
            "consolidated_rows": [{"feature_row": "docx-3", "into_feature_row": "docx-1", "reason": "Grouped with Base."}],
            "status_changes": [{"feature_row": "docx-1", "current_status": {"status": "🟡", "read": "✅", "write": "🟡"}, "reason": "Bounded profile."}],
            "additional_rows": [{"id": "docx-current-added", "format": "docx", "source_matrix": "matrix.md",
                                 "feature": "Added", "current_status": {"status": "✅", "read": "✅", "write": "✅"},
                                 "scope_limit": "Added feature without historical evidence."}],
        }

    def test_reconciliation_retains_distinct_current_and_historical_identities(self) -> None:
        reconciled = reconcile_matrices(self.inventory, self.document)
        self.assertEqual([(row["id"], row["historical_feature_row"]) for row in reconciled["rows"]], [
            ("docx-1", "docx-1"), ("docx-2", "docx-2"), ("docx-current-added", None),
        ])
        self.assertEqual(self.inventory[0]["matrix_status"], "✅")
        self.assertEqual(self.inventory[1]["feature"], "Old name")
        self.assertEqual(self.inventory[2]["feature"], "Consolidated")

    def test_unreviewed_matrix_drift_fails_closed(self) -> None:
        original = self.matrix.read_text()
        mutations = [
            original + "| Unaccounted | ✅ | ✅ | ✅ |\n",
            original.replace("| Base | 🟡 | ✅ | 🟡 |\n", ""),
            original.replace("| Base |", "| New base name |"),
            original.replace("| Base | 🟡 | ✅ | 🟡 |", "| Base | ✅ | ✅ | ✅ |"),
            original + "| Base | 🟡 | ✅ | 🟡 |\n",
            original + "| Consolidated | ✅ | ✅ | ✅ |\n",
            original.replace("| Added | ✅ | ✅ | ✅ |\n", ""),
            "| Feature | Read | Write |\n| --- | --- | --- |\n| Base | ✅ | ✅ |\n",
        ]
        for changed in mutations:
            with self.subTest(matrix=changed):
                self.matrix.write_text(changed)
                with self.assertRaises(ValueError):
                    reconcile_matrices(self.inventory, self.document)

    def test_invalid_or_stale_reconciliation_entries_are_rejected(self) -> None:
        for case in ("unknown", "duplicate", "blank_reason", "redundant_rename", "redundant_status",
                     "removed_target", "wrong_format", "duplicate_added_id", "duplicate_binding", "blank_scope"):
            with self.subTest(case=case):
                candidate = copy.deepcopy(self.document)
                if case == "unknown":
                    candidate["renamed_rows"][0]["feature_row"] = "docx-unknown"
                elif case == "duplicate":
                    candidate["status_changes"] *= 2
                elif case == "blank_reason":
                    candidate["consolidated_rows"][0]["reason"] = " "
                elif case == "redundant_rename":
                    candidate["renamed_rows"][0]["current_feature"] = "Old name"
                elif case == "redundant_status":
                    candidate["status_changes"][0]["current_status"] = {"status": "✅", "read": "✅", "write": "✅"}
                elif case == "removed_target":
                    candidate["consolidated_rows"][0]["into_feature_row"] = "docx-3"
                elif case == "wrong_format":
                    candidate["additional_rows"][0]["format"] = "xlsb"
                elif case == "duplicate_added_id":
                    candidate["additional_rows"] *= 2
                elif case == "duplicate_binding":
                    candidate["additional_rows"][0]["feature"] = "Renamed"
                else:
                    candidate["additional_rows"][0]["scope_limit"] = " "
                with self.assertRaises(ValueError):
                    reconcile_matrices(self.inventory, candidate)

    def test_checklist_status_cells_reject_narrative_and_malformed_tables(self) -> None:
        original = self.matrix.read_text()
        for invalid in (
            original.replace("| Base | 🟡", "| Base | 🟡 implementation details"),
            original.replace("| Read |", "| Test results |"),
            original.replace("| --- | --- | --- | --- |\n", ""),
            original.replace("| Added | ✅ | ✅ | ✅ |", "| Added | ✅ | ✅ |"),
            "| Feature | Read | Write |\n",
            "No feature table.\n",
        ):
            with self.subTest(matrix=invalid), self.assertRaises(ValueError):
                matrix_rows(invalid, "matrix.md")


if __name__ == "__main__":
    unittest.main()
