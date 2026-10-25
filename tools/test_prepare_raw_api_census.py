#!/usr/bin/env python3
"""Source-only tests for the deterministic current raw-census matrix."""

from __future__ import annotations

import json
from pathlib import Path
import subprocess
import unittest
from unittest.mock import patch

from tools import prepare_raw_api_census as census


ROOT = Path(__file__).resolve().parents[1]


class RawApiCensusPreparationTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.matrix = census.build_matrix(ROOT)

    def test_current_matrix_has_71_glossary_and_13_package_only_profiles(self) -> None:
        self.assertEqual(
            self.matrix["counts"],
            {
                "direct_packages": 68,
                "glossary_profiles": 71,
                "package_only_profiles": 13,
                "profiles": 84,
            },
        )
        self.assertEqual(self.matrix["state"], "incomplete")
        self.assertEqual(self.matrix["entries"], [])
        self.assertEqual(self.matrix["lanes"], [])
        self.assertEqual(self.matrix["captures"], [])

    def test_glossary_profiles_preserve_current_tuples_and_order(self) -> None:
        glossary = json.loads((ROOT / census.GLOSSARY_PATH).read_text(encoding="utf-8"))
        actual = self.matrix["profiles"][: len(glossary["profiles"])]
        for expected, observed in zip(glossary["profiles"], actual):
            self.assertEqual(observed["id"], expected["id"])
            self.assertEqual(observed["package"], expected["package"])
            self.assertEqual(observed["default_features"], expected["default_features"])
            self.assertEqual(observed["features"], expected["features"])
            self.assertEqual(observed["origin"], "glossary")

    def test_package_only_profiles_are_the_direct_non_glossary_set(self) -> None:
        only = self.matrix["profiles"][71:]
        self.assertEqual(
            [profile["package"] for profile in only],
            [
                "litchi-cli",
                "litchi-codepage",
                "litchi-eval",
                "litchi-iwa-cache",
                "litchi-iwa-package",
                "litchi-iwa-text",
                "litchi-odf",
                "litchi-py",
                "litchi-sheet",
                "litchi-slide",
                "litchi-word",
                "xml-minifier",
                "xml-minifier-macros",
            ],
        )
        self.assertEqual(only[2]["features"], ["web_functions"])
        self.assertEqual(only[6]["features"], ["all"])
        self.assertTrue(all(profile["default_features"] is False for profile in only))
        self.assertTrue(all(profile["origin"] == "package_only" for profile in only))

    def test_source_and_manifest_identities_are_bound_to_pinned_commit(self) -> None:
        source = self.matrix["source"]
        self.assertEqual(len(source["commit"]), 40)
        subprocess.run(
            ["git", "-C", str(ROOT), "cat-file", "-e", f"{source['commit']}^{{commit}}"],
            check=True,
        )
        for identity in (source["glossary"], source["workspace"]):
            self.assertEqual(identity["blob"], census._git_blob(ROOT, identity["path"]))
            self.assertEqual(identity["blob"], census._git_blob_at(ROOT, source["commit"], identity["path"]))
            self.assertGreater(identity["bytes"], 0)
            self.assertEqual(
                identity["sha256"],
                census._sha256((ROOT / identity["path"]).read_bytes()),
            )
        self.assertEqual(self.matrix["legacy_coverage"]["profiles"], 15)
        self.assertEqual(self.matrix["legacy_coverage"]["entries"], 22)
        self.assertEqual(self.matrix["legacy_coverage"]["lanes"], 12)

    def test_dirty_input_overlay_is_rejected(self) -> None:
        with patch.object(census, "_worktree_blob", return_value="0" * 40):
            with self.assertRaisesRegex(census.CensusError, "working-tree bytes"):
                census._file_identity(ROOT, census.GLOSSARY_PATH)

    def test_legacy_manifest_must_match_the_selected_source_commit(self) -> None:
        source_commit = self.matrix["source"]["commit"]
        original = census._git_blob_at

        def hide_legacy_blob(root: Path, revision: str, relative: str) -> str:
            if revision == source_commit and relative == census.LEGACY_MANIFEST_PATH:
                return "0" * 40
            return original(root, revision, relative)

        with patch.object(census, "_git_blob_at", side_effect=hide_legacy_blob):
            with self.assertRaisesRegex(census.CensusError, "raw_api_coverage.json differs"):
                census.build_matrix(ROOT, source_commit=source_commit)

    def test_matrix_carries_approved_f5_policy_without_capture_claims(self) -> None:
        policy = self.matrix["capture_policy"]
        self.assertIn("run_safe_gate.py", policy["runner_path"])
        self.assertIn("/bin/cargo", policy["cap_wrapper_path"])
        self.assertEqual(policy["guards"]["build_jobs"], 1)
        self.assertEqual(policy["guards"]["memory_max"], "20G")
        self.assertIn("rustdoc inputs", self.matrix["reason"])

    def test_checked_in_matrix_is_deterministic(self) -> None:
        output = ROOT / "tools/raw_api_census_matrix.json"
        stored = json.loads(output.read_text(encoding="utf-8"))
        pinned = census.build_matrix(ROOT, source_commit=stored["source"]["commit"])
        expected = json.dumps(pinned, indent=2, sort_keys=True) + "\n"
        self.assertEqual(output.read_text(encoding="utf-8"), expected)
        self.assertFalse(census.write_matrix(pinned, output, check=True))


if __name__ == "__main__":
    unittest.main()
