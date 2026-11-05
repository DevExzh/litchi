"""Integrity of the native Office corpus; this is not semantic API coverage.

Run with ``python3 -m unittest tools.test_native_office_corpus``. The native
authoring inventory remains historical evidence; Rust regression tests must
separately establish the behavior of the current library.
"""

from collections import Counter
import hashlib
import json
from pathlib import Path
import unittest
import zipfile

from tools.native_office_coverage import report


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


if __name__ == "__main__":
    unittest.main()
