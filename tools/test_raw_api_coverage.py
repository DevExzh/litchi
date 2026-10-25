from __future__ import annotations

import json
from pathlib import Path
import tempfile
import unittest

from tools import raw_api_coverage as coverage


def rustdoc_fixture() -> dict:
    """One raw callable and one explicitly semantic callable."""
    return {
        "root": 0,
        "includes_private": False,
        "index": {
            "0": {
                "name": "fixture",
                "crate_id": 0,
                "visibility": "public",
                "inner": {"module": {"items": [1, 2, 3], "is_crate": True}},
            },
            "1": {
                "name": "SourceLimits",
                "crate_id": 0,
                "visibility": "public",
                "inner": {"struct": {"fields": [], "impls": []}},
            },
            "2": {
                "name": "parse",
                "crate_id": 0,
                "visibility": "public",
                "inner": {"function": {
                    "sig": {"inputs": [
                        ["data", {"primitive": "unit"}],
                        ["limits", {"borrowed_ref": {
                            "is_mutable": False,
                            "type": {"resolved_path": {"id": 1, "args": None}},
                        }}],
                    ], "output": {"primitive": "unit"}},
                    "generics": {"params": []},
                }},
            },
            "3": {
                "name": "describe",
                "crate_id": 0,
                "visibility": "public",
                "inner": {"function": {
                    "sig": {"inputs": [], "output": {"primitive": "unit"}},
                    "generics": {"params": []},
                }},
            },
        },
        "paths": {
            "0": {"path": ["fixture"], "kind": "module", "crate_id": 0},
            "1": {"path": ["fixture", "SourceLimits"], "kind": "struct", "crate_id": 0},
            "2": {"path": ["fixture", "parse"], "kind": "function", "crate_id": 0},
            "3": {"path": ["fixture", "describe"], "kind": "function", "crate_id": 0},
        },
    }


def glossary_fixture(path: Path) -> None:
    path.write_text(json.dumps({
        "profiles": [{
            "id": "fixture-profile",
            "package": "fixture",
            "default_features": True,
            "features": [],
        }],
    }), encoding="utf-8")


def entry(path: str, kind: str, disposition: str, *, signature=None) -> dict:
    value = {
        "profile": "fixture-profile",
        "path": path,
        "kind": kind,
        "disposition": disposition,
        "source": {
            "path": "crates/fixture/src/lib.rs",
            "blob": "a" * 40,
            "line": 1,
            "symbol": path.rsplit("::", 1)[-1],
            "symbol_kind": "rust_item",
        },
        "evidence": "synthetic rustdoc fixture evidence",
    }
    if signature is not None:
        value["required_signature"] = signature
    return value


def manifest(entries: list[dict], *, state: str = "complete") -> dict:
    value = {
        "schema_version": 1,
        "source_commit": "b" * 40,
        "coverage": {
            "state": state,
            "reason": "synthetic coverage fixture",
            "expected_profile_source": "tools/api_glossary.json",
            "expected_workspace_source": "Cargo.toml",
        },
        "profiles": [{
            "id": "fixture-profile",
            "package": "fixture",
            "default_features": True,
            "features": [],
        }],
        "entries": entries,
        "lanes": [],
    }
    for item in entries:
        if item["disposition"] != "active_lane":
            continue
        value["lanes"].append({
            "id": item["lane"],
            "package": "fixture",
            "scope": "fixture",
            "disposition": "active_lane",
            "source": {
                "path": "crates/fixture/src/lib.rs",
                "blob": "a" * 40,
                "line": 1,
                "symbol": "lane-anchor",
                "symbol_kind": "descriptive",
            },
            "evidence": "synthetic active lane",
        })
    return value


class RawApiCoverageTests(unittest.TestCase):
    def setUp(self) -> None:
        self.tmp = tempfile.TemporaryDirectory()
        self.root = Path(self.tmp.name)
        self.glossary = self.root / "glossary.json"
        glossary_fixture(self.glossary)

    def tearDown(self) -> None:
        self.tmp.cleanup()

    def loaded(self, value: dict, *, require_complete: bool = False,
               extra_rustdoc: bool = False) -> dict:
        manifest_path = self.root / "manifest.json"
        manifest_path.write_text(json.dumps(value), encoding="utf-8")
        loaded = coverage.load_manifest(manifest_path)
        rustdoc = self.root / "fixture.json"
        rustdoc.write_text(json.dumps(rustdoc_fixture()), encoding="utf-8")
        bindings = {"fixture-profile": rustdoc}
        if extra_rustdoc:
            bindings["extra-profile"] = rustdoc
        return coverage.validate_manifest(
            loaded, root=self.root, glossary_path=self.glossary,
            rustdoc=bindings, require_complete=require_complete,
            verify_sources=False,
        )

    @staticmethod
    def raw_signature(limits_type: str = "fixture::SourceLimits") -> dict:
        return {
            "inputs": ["data", "limits"],
            "limits_argument": "limits",
            "limits_type": limits_type,
        }

    def test_exact_census_accepts_registered_raw_signature_and_semantic_anchor(self) -> None:
        value = manifest([
            entry("fixture::parse", "raw_function", "required_raw_limits",
                  signature=self.raw_signature()),
            entry("fixture::describe", "semantic_method", "semantic_value"),
        ])
        report = self.loaded(value, require_complete=True)
        self.assertTrue(report["complete"], report)
        self.assertEqual(report["callable_counts"], {"fixture-profile": 2})
        self.assertEqual(report["unclassified"], [])

    def test_missing_callable_is_unclassified_even_when_name_is_not_raw(self) -> None:
        value = manifest([entry(
            "fixture::parse", "raw_function", "required_raw_limits",
            signature=self.raw_signature(),
        )])
        report = self.loaded(value)
        self.assertEqual(report["unclassified"], ["fixture-profile::fixture::describe"])
        self.assertFalse(report["complete"])

    def test_raw_signature_type_mismatch_fails_closed(self) -> None:
        value = manifest([
            entry("fixture::parse", "raw_function", "required_raw_limits",
                  signature=self.raw_signature("fixture::WrongLimits")),
            entry("fixture::describe", "semantic_method", "semantic_value"),
        ])
        report = self.loaded(value)
        self.assertEqual(len(report["signature_failures"]), 1)
        self.assertIn("limits type identity", report["signature_failures"][0])

    def test_active_lane_and_incomplete_state_fail_final_mode(self) -> None:
        active = entry("fixture::describe", "semantic_method", "active_lane")
        active["lane"] = "0362ul-test"
        value = manifest([active])
        value["lanes"] = []
        value["lanes"].append({
            "id": "0362ul-test", "package": "fixture", "scope": "fixture",
            "disposition": "active_lane",
            "source": {
                "path": "crates/fixture/src/lib.rs", "blob": "a" * 40,
                "line": 1, "symbol": "lane-anchor", "symbol_kind": "descriptive",
            },
            "evidence": "synthetic active lane",
        })
        value["coverage"]["state"] = "incomplete"
        report = self.loaded(value, require_complete=True)
        self.assertFalse(report["complete"])
        self.assertEqual(report["active_entries"], ["fixture-profile::fixture::describe"])
        self.assertIn("failure", report)

    def test_private_removed_is_stale_if_rustdoc_still_exports_it(self) -> None:
        value = manifest([
            entry("fixture::parse", "raw_function", "required_raw_limits",
                  signature=self.raw_signature()),
            entry("fixture::describe", "removed_callable", "private_removed"),
        ])
        report = self.loaded(value)
        self.assertEqual(report["stale_entries"], [
            "fixture-profile::fixture::describe: private_removed item remains public",
        ])

    def test_manifest_rejects_generic_disposition_and_unknown_fields(self) -> None:
        value = manifest([entry("fixture::parse", "raw_function", "out_of_scope")])
        with self.assertRaises(coverage.CoverageError):
            path = self.root / "manifest.json"
            path.write_text(json.dumps(value), encoding="utf-8")
            coverage.load_manifest(path)

    def test_extra_rustdoc_binding_cannot_make_final_mode_pass(self) -> None:
        value = manifest([
            entry("fixture::parse", "raw_function", "required_raw_limits",
                  signature=self.raw_signature()),
            entry("fixture::describe", "semantic_method", "semantic_value"),
        ])
        report = self.loaded(value, require_complete=True, extra_rustdoc=True)
        self.assertEqual(report["extra_inputs"], ["extra-profile"])
        self.assertFalse(report["complete"])

    def test_undeclared_active_lane_is_rejected(self) -> None:
        active = entry("fixture::describe", "semantic_method", "active_lane")
        active["lane"] = "0362ul-missing"
        value = manifest([active])
        value["lanes"] = []
        path = self.root / "manifest.json"
        path.write_text(json.dumps(value), encoding="utf-8")
        with self.assertRaises(coverage.CoverageError):
            coverage.load_manifest(path)

    def test_callable_source_anchor_must_be_rust_item(self) -> None:
        value = manifest([entry("fixture::describe", "semantic_method", "semantic_value")])
        value["entries"][0]["source"]["symbol_kind"] = "descriptive"
        value["entries"][0]["source"]["symbol"] = "arbitrary-label"
        path = self.root / "manifest.json"
        path.write_text(json.dumps(value), encoding="utf-8")
        with self.assertRaisesRegex(coverage.CoverageError, "rust_item"):
            coverage.load_manifest(path)

    def test_source_anchor_requires_blob_bounds_and_verified_symbol(self) -> None:
        root = Path(__file__).resolve().parents[1]
        head = coverage._git_head(root)
        self.assertIsNotNone(head)
        blob = coverage._git_blob(root, head, "tools/check_api_glossary.py")
        self.assertIsNotNone(blob)
        value = manifest([])
        value["source_commit"] = head
        value["lanes"] = [{
            "id": "test-source", "package": "fixture", "scope": "fixture",
            "disposition": "active_lane",
            "source": {
                "path": "tools/check_api_glossary.py", "blob": blob,
                "line": 1, "symbol": "module", "symbol_kind": "descriptive",
            },
            "evidence": "source identity test",
        }]
        path = self.root / "manifest.json"
        path.write_text(json.dumps(value), encoding="utf-8")
        loaded = coverage.load_manifest(path)
        self.assertEqual(coverage.verify_source_anchors(loaded, root), [])

        tree = coverage._git_object(root, head, "tools")
        self.assertIsNotNone(tree)
        loaded["source_commit"] = tree[0]
        self.assertIn("does not resolve to a Git commit", coverage.verify_source_anchors(loaded, root)[0])
        loaded["source_commit"] = head

        loaded["lanes"][0]["source"]["blob"] = "c" * 40
        self.assertIn("blob", coverage.verify_source_anchors(loaded, root)[0])
        loaded["lanes"][0]["source"]["blob"] = blob
        loaded["lanes"][0]["source"]["line"] = 999999
        self.assertIn("exceeds", coverage.verify_source_anchors(loaded, root)[0])

        loaded["lanes"][0]["source"]["line"] = 1
        loaded["lanes"][0]["source"]["symbol"] = "not_the_function"
        loaded["lanes"][0]["source"]["symbol_kind"] = "rust_item"
        self.assertIn("does not contain rust item", coverage.verify_source_anchors(loaded, root)[0])

        loaded["lanes"][0]["source"]["path"] = "tools"
        loaded["lanes"][0]["source"]["blob"] = tree[0]
        loaded["lanes"][0]["source"]["symbol_kind"] = "descriptive"
        self.assertIn("not a blob", coverage.verify_source_anchors(loaded, root)[0])

    def test_source_anchor_schema_requires_symbol_kind(self) -> None:
        value = manifest([entry("fixture::describe", "semantic_method", "semantic_value")])
        value["entries"][0]["source"].pop("symbol_kind")
        path = self.root / "manifest.json"
        path.write_text(json.dumps(value), encoding="utf-8")
        with self.assertRaises(coverage.CoverageError):
            coverage.load_manifest(path)


if __name__ == "__main__":
    unittest.main()
