"""Independent corpus checks against retained native fixtures and tampering."""

from __future__ import annotations

import copy
import hashlib
import json
from pathlib import Path
import sys
import unittest
import zipfile

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
import corpus_oracle


class CorpusOracleTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.report = json.loads((HERE / "consumer/dev56/cli-report.json").read_text())

    def test_retained_report_and_native_archive_main_bytes(self):
        for case in self.report["cases"]:
            with self.subTest(text=case["text_mode"]):
                corpus_oracle.validate_case(case)
                expected = corpus_oracle.expected_case(
                    case["source_count"], case["authored_count"],
                    case["chunk_mode"], case["text_mode"],
                )
                directory = HERE / "consumer/dev56" / case["text_mode"]
                for role in ("source", "candidate"):
                    with zipfile.ZipFile(directory / f"{role}.docx") as archive:
                        xml = archive.read("word/document.xml")
                    record = expected["source" if role == "source" else "oracle"]
                    prefix = "" if role == "source" else "candidate_"
                    self.assertEqual(len(xml), record[f"{prefix}main_xml_bytes"])
                    self.assertEqual(hashlib.sha256(xml).hexdigest(), record[f"{prefix}main_xml_sha256"])

    def test_self_consistent_authored_hash_corruption_is_rejected(self):
        case = copy.deepcopy(self.report["cases"][0])
        for section in (case["authored"], case["proof"]["authored"]):
            section["expected_event_sha256"] = "0" * 64
        with self.assertRaisesRegex(ValueError, "independent generator"):
            corpus_oracle.validate_case(case)

    def test_self_consistent_candidate_semantic_corruption_is_rejected(self):
        case = copy.deepcopy(self.report["cases"][0])
        case["oracle"]["candidate_semantic"]["text_sha256"] = "a" * 64
        with self.assertRaisesRegex(ValueError, "candidate_semantic"):
            corpus_oracle.validate_case(case)

    def test_chunk_partition_changes_event_hash_but_preserves_encoded_xml(self):
        one = corpus_oracle.expected_case(64, 64, "one", "near_limit")
        chunks = corpus_oracle.expected_case(64, 64, "fixed64", "near_limit")
        self.assertNotEqual(one["authored"]["expected_event_sha256"], chunks["authored"]["expected_event_sha256"])
        self.assertGreater(chunks["authored"]["event_count"], one["authored"]["event_count"])
        self.assertEqual(one["oracle"], chunks["oracle"])
        self.assertEqual(one["authored"]["expected_encoded_sha256"], chunks["authored"]["expected_encoded_sha256"])

    def test_finite_domain_and_cache_cannot_be_mutated_by_caller(self):
        for args in ((True, 64, "one", "short"), (1, 0, "one", "short"),
                     (131_073, 64, "one", "short"), (64, 16_384, "one", "near_limit"),
                     ([64], 64, "one", "short"), (64, 64, ["one"], "short")):
            with self.subTest(args=args), self.assertRaises(ValueError):
                corpus_oracle.expected_case(*args)
        record = corpus_oracle.expected_case(64, 64, "fixed64", "short")
        record["authored"]["text_bytes"] = 0
        self.assertGreater(corpus_oracle.expected_case(64, 64, "fixed64", "short")["authored"]["text_bytes"], 0)

    def test_booleans_cannot_replace_zero_integer_evidence(self):
        expected = corpus_oracle.expected_case(64, 64, "fixed64", "empty")
        case = dict(expected, source_count=64, authored_count=64,
                    chunk_mode="fixed64", text_mode="empty")
        case["proof"]["authored"] = copy.deepcopy(case["authored"])
        corpus_oracle.validate_case(case)
        for section in ("authored", "proof"):
            for field in ("text_bytes", "xml_entity_reference_count"):
                changed = copy.deepcopy(case)
                record = changed[section] if section == "authored" else changed["proof"]["authored"]
                record[field] = False
                with self.subTest(section=section, field=field), self.assertRaises(ValueError):
                    corpus_oracle.validate_case(changed)


if __name__ == "__main__":
    unittest.main()
