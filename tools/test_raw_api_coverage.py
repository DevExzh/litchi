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
    return value


class RawApiCoverageTests(unittest.TestCase):
    def setUp(self) -> None:
        self.tmp = tempfile.TemporaryDirectory()
        self.root = Path(self.tmp.name)
        self.glossary = self.root / "glossary.json"
        glossary_fixture(self.glossary)

    def tearDown(self) -> None:
        self.tmp.cleanup()

    def loaded(self, value: dict, *, require_complete: bool = False) -> dict:
        path = self.root / "manifest.json"
        path.write_text(json.dumps(value), encoding="utf-8")
        loaded = coverage.load_manifest(path)
        rustdoc = self.root / "fixture.json"
        rustdoc.write_text(json.dumps(rustdoc_fixture()), encoding="utf-8")
        return coverage.validate_manifest(
            loaded, root=self.root, glossary_path=self.glossary,
            rustdoc={"fixture-profile": rustdoc},
            require_complete=require_complete, verify_sources=False,
        )

    def test_exact_census_accepts_registered_raw_signature_and_semantic_anchor(self) -> None:
        value = manifest([
            entry(
                "fixture::parse", "raw_function", "required_raw_limits",
                signature={
                    "inputs": ["data", "limits"],
                    "limits_argument": "limits",
                    "limits_type": "fixture::SourceLimits",
                },
            ),
            entry("fixture::describe", "semantic_method", "semantic_value"),
        ])
        report = self.loaded(value, require_complete=True)
        self.assertTrue(report["complete"], report)
        self.assertEqual(report["callable_counts"], {"fixture-profile": 2})
        self.assertEqual(report["unclassified"], [])

    def test_missing_callable_is_unclassified_even_when_name_is_not_raw(self) -> None:
        value = manifest([entry(
            "fixture::parse", "raw_function", "required_raw_limits",
            signature={
                "inputs": ["data", "limits"],
                "limits_argument": "limits",
                "limits_type": "fixture::SourceLimits",
            },
        )])
        report = self.loaded(value)
        self.assertEqual(report["unclassified"], ["fixture-profile::fixture::describe"])
        self.assertFalse(report["complete"])

    def test_raw_signature_type_mismatch_fails_closed(self) -> None:
        value = manifest([
            entry(
                "fixture::parse", "raw_function", "required_raw_limits",
                signature={
                    "inputs": ["data", "limits"],
                    "limits_argument": "limits",
                    "limits_type": "fixture::WrongLimits",
                },
            ),
            entry("fixture::describe", "semantic_method", "semantic_value"),
        ])
        report = self.loaded(value)
        self.assertEqual(len(report["signature_failures"]), 1)
        self.assertIn("limits type identity", report["signature_failures"][0])

    def test_active_lane_and_incomplete_state_fail_final_mode(self) -> None:
        active = entry("fixture::describe", "raw_function", "active_lane")
        active["lane"] = "0362ul-test"
        value = manifest([active], state="incomplete")
        report = self.loaded(value, require_complete=True)
        self.assertFalse(report["complete"])
        self.assertEqual(report["active_entries"], ["fixture-profile::fixture::describe"])
        self.assertIn("failure", report)

    def test_private_removed_is_stale_if_rustdoc_still_exports_it(self) -> None:
        removed = entry("fixture::describe", "removed_callable", "private_removed")
        value = manifest([
            entry(
                "fixture::parse", "raw_function", "required_raw_limits",
                signature={
                    "inputs": ["data", "limits"],
                    "limits_argument": "limits",
                    "limits_type": "fixture::SourceLimits",
                },
            ),
            removed,
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

    def test_source_identity_checker_detects_tampering(self) -> None:
        # The repository source check is exercised against this checkout's
        # immutable HEAD; the manifest fixture remains source-disabled above
        # so the callable tests do not depend on a complete checkout.
        root = Path(__file__).resolve().parents[1]
        head = coverage._git_head(root)
        self.assertIsNotNone(head)
        blob = coverage._git_blob(root, head, "tools/check_api_glossary.py")
        self.assertIsNotNone(blob)
        value = manifest([])
        value["source_commit"] = head
        value["lanes"] = [{
            "id": "test-source",
            "package": "fixture",
            "scope": "fixture",
            "disposition": "active_lane",
            "source": {
                "path": "tools/check_api_glossary.py",
                "blob": blob,
                "line": 1,
                "symbol": "module",
            },
            "evidence": "source identity test",
        }]
        path = self.root / "manifest.json"
        path.write_text(json.dumps(value), encoding="utf-8")
        loaded = coverage.load_manifest(path)
        failures = coverage.verify_source_anchors(loaded, root)
        self.assertEqual(failures, [])
        loaded["lanes"][0]["source"]["blob"] = "c" * 40
        failures = coverage.verify_source_anchors(loaded, root)
        self.assertEqual(len(failures), 1)
        self.assertIn("source tools/check_api_glossary.py blob", failures[0])


if __name__ == "__main__":
    unittest.main()
