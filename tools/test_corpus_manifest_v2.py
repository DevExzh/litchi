#!/usr/bin/env python3
"""Pure-Python checks for the additive corpus-manifest migration."""

from __future__ import annotations

import copy
import hashlib
import json
import unittest
from pathlib import Path

from tools.generate_corpus_manifest_v2 import canonical_bytes, generate


ROOT = Path(__file__).resolve().parents[1]
V1_PATH = ROOT / "docs/performance/results/perf-regression-default-manifest-v1.json"
V2_PATH = ROOT / "docs/performance/results/perf-corpus-manifest-v2.json"
SCHEMA_PATH = ROOT / "docs/performance/schemas/corpus-manifest-v2.schema.json"


def sha256(value: bytes) -> str:
    return hashlib.sha256(value).hexdigest()


def content_set(catalog: dict) -> dict:
    return {
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


class CorpusManifestV2Tests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.v1 = json.loads(V1_PATH.read_text(encoding="utf-8"))
        cls.v2 = json.loads(V2_PATH.read_text(encoding="utf-8"))

    def test_checked_schema_is_valid_json(self):
        schema = json.loads(SCHEMA_PATH.read_text(encoding="utf-8"))
        self.assertEqual(schema["$schema"], "https://json-schema.org/draft/2020-12/schema")
        self.assertEqual(schema["properties"]["manifest_version"]["const"], 2)

    def test_v1_identity_is_unchanged(self):
        self.assertEqual(self.v1["schema_version"], 1)
        self.assertEqual(self.v1["manifest_kind"], "case-corpus-key-identity")
        self.assertEqual(self.v1["result_keys_sha256"],
                         "3b57c3b5aef77f5149d520fd885194d1fd8734460b28bff9d317d1cd840c246f")

    def test_catalog_shape_and_bindings(self):
        self.assertEqual(self.v2["manifest_version"], 2)
        self.assertEqual(self.v2["manifest_kind"], "corpus-catalog")
        self.assertEqual(len(self.v2["corpora"]), 28)
        self.assertEqual(len(self.v2["case_bindings"]), self.v1["result_count"])
        ids = [corpus["id"] for corpus in self.v2["corpora"]]
        self.assertEqual(ids, sorted(ids))
        self.assertEqual(len(ids), len(set(ids)))
        keys = [
            (binding["case"], binding["corpus_id"])
            for binding in self.v2["case_bindings"]
        ]
        self.assertEqual(keys, sorted(keys))
        self.assertEqual(len(keys), len(set(keys)))

    def test_catalog_hashes_are_reproducible(self):
        without_hash = copy.deepcopy(self.v2)
        without_hash.pop("catalog_sha256")
        self.assertEqual(
            self.v2["catalog_sha256"], sha256(canonical_bytes(without_hash))
        )
        self.assertEqual(
            self.v2["content_set_sha256"],
            sha256(canonical_bytes(content_set(self.v2))),
        )

    def test_generator_reproduces_checked_catalog(self):
        generated = generate(self.v1, self.v2["build"]["git_revision"])
        self.assertEqual(generated, self.v2)

    def test_migration_keeps_unknowns_explicit(self):
        for corpus in self.v2["corpora"]:
            self.assertEqual(corpus["legacy_v1"], self.v1["corpora"][corpus["name"]])
            self.assertEqual(corpus["members"]["status"], "unavailable")
            self.assertEqual(corpus["relationships"]["status"], "unknown")
            self.assertEqual(corpus["input"]["validity"], "unknown")
            self.assertIsNone(corpus["limits"]["profile_id"])
            self.assertEqual(corpus["security"]["encryption"]["state"], "unknown")
            self.assertEqual(corpus["security"]["signature"]["state"], "unknown")
            self.assertEqual(corpus["security"]["macros"]["state"], "unknown")

    def test_adding_reference_does_not_change_v1_result_identity(self):
        original = copy.deepcopy(self.v1)
        additive = copy.deepcopy(original)
        additive["corpus_catalog"] = {
            "manifest_version": self.v2["manifest_version"],
            "catalog_id": self.v2["catalog_id"],
            "catalog_sha256": self.v2["catalog_sha256"],
            "content_set_sha256": self.v2["content_set_sha256"],
        }
        self.assertEqual(
            original["result_keys_sha256"], additive["result_keys_sha256"]
        )


if __name__ == "__main__":
    unittest.main()
