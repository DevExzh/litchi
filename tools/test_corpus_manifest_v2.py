#!/usr/bin/env python3
"""Pure-Python checks for the additive corpus-manifest migration."""

from __future__ import annotations

import copy
import hashlib
import json
import unittest
from pathlib import Path

from tools.generate_corpus_manifest_v2 import FAMILY_MAP, canonical_bytes, generate


ROOT = Path(__file__).resolve().parents[1]
V1_PATH = ROOT / "docs/performance/results/perf-regression-default-manifest-v1.json"
V2_PATH = ROOT / "docs/performance/results/perf-corpus-manifest-v2.json"
SCHEMA_PATH = ROOT / "docs/performance/schemas/corpus-manifest-v2.schema.json"
EXPECTED_RESULT_KEYS_SHA256 = (
    "3b57c3b5aef77f5149d520fd885194d1fd8734460b28bff9d317d1cd840c246f"
)
EXPECTED_CONTENT_SET_SHA256 = (
    "f0fad22aa4deed6ccf2f8ecc202222ee2c901aec3fe908e69edee0dbd46b1d71"
)


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
        self.assertEqual(self.v1["result_keys_sha256"], EXPECTED_RESULT_KEYS_SHA256)

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
        self.assertEqual(self.v2["content_set_sha256"], EXPECTED_CONTENT_SET_SHA256)

    def test_checked_default_uses_the_documented_family_map(self):
        expected_families = {
            "litchi-cfb-synthetic-v1": (
                "cfb",
                "litchi-perf.cfb-payload-v1",
                "indexed-formula-v1",
            ),
            "litchi-opc-synthetic-v2": (
                "opc",
                "litchi-perf.opc-payload-v1",
                "indexed-formula-v1",
            ),
            "litchi-legacy-writer-v1": (
                "legacy-writer",
                "litchi-perf.legacy-writer-v1",
                "none",
            ),
            "litchi-xlsx-synthetic-v1": (
                "xlsx",
                "litchi-perf.xlsx-integer-grid-v1",
                "none",
            ),
        }
        self.assertEqual(set(FAMILY_MAP), set(expected_families))
        self.assertEqual(
            {
                generator: (
                    metadata["family"],
                    metadata["algorithm_id"],
                    metadata["seed_spec"],
                )
                for generator, metadata in FAMILY_MAP.items()
            },
            expected_families,
        )
        seen = set()
        for corpus in self.v2["corpora"]:
            seen.add(corpus["generator"]["id"])
            family, algorithm_id, seed_spec = expected_families[
                corpus["generator"]["id"]
            ]
            generator = corpus["generator"]
            self.assertEqual(generator["kind"], "synthetic")
            self.assertEqual(generator["algorithm_id"], algorithm_id)
            self.assertEqual(generator["seed_spec"], seed_spec)
            self.assertEqual(generator["parameters"]["family"], family)
            self.assertEqual(
                corpus["provenance"]["source_path"],
                "tools/perf-baseline/src/lib.rs",
            )
            self.assertEqual(corpus["provenance"]["source_kind"], "generated")
            self.assertIsNone(corpus["provenance"]["source_sha256"])
        self.assertEqual(seen, set(expected_families))

    def test_seed_and_formula_metadata_is_truthful(self):
        for corpus in self.v2["corpora"]:
            generator = corpus["generator"]
            parameters = generator["parameters"]
            if generator["id"] in {
                "litchi-cfb-synthetic-v1",
                "litchi-opc-synthetic-v2",
            }:
                self.assertEqual(generator["seed_spec"], "indexed-formula-v1")
                self.assertIn("index", parameters["payload_formula"])
                if corpus["legacy_v1"]["payload_kind"] == "incompressible":
                    self.assertIn("xorshift64", parameters["payload_formula"])
                    self.assertIn("0x9e3779b97f4a7c15", parameters["payload_formula"])
                else:
                    self.assertEqual(parameters["payload_block_bytes"], 45)
            elif generator["id"] == "litchi-xlsx-synthetic-v1":
                self.assertEqual(generator["seed_spec"], "none")
                self.assertIn("sheet * 1_000_000", parameters["value_formula"])
                self.assertEqual(
                    parameters["one_percent_update_count"],
                    corpus["legacy_v1"]["xlsx"]["one_percent_update_count"],
                )
            else:
                self.assertEqual(generator["id"], "litchi-legacy-writer-v1")
                self.assertEqual(generator["seed_spec"], "none")
                self.assertIn("writer_text_template", parameters)
                self.assertIn("payload_heavy_repeat_block", parameters)

    def test_known_metadata_does_not_fill_unavailable_archive_facts(self):
        for corpus in self.v2["corpora"]:
            self.assertEqual(corpus["members"], {"status": "unavailable", "items": []})
            self.assertEqual(corpus["relationships"]["status"], "unknown")
            self.assertIsNone(corpus["relationships"]["relationship_count"])
            self.assertIsNone(corpus["limits"]["profile_id"])
            self.assertIsNone(corpus["limits"]["profile_sha256"])
            self.assertIsNone(corpus["provenance"]["source_sha256"])

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
