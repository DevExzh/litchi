from __future__ import annotations

import unittest

from tools import api_debt_inventory as inventory


def owner_fixture() -> dict:
    """A small rustdoc graph with a non-heuristic owner and public aliases."""
    index = {}
    paths = {}

    def add(item_id, name, path, kind, inner, *, visibility="public", span=False):
        key = str(item_id)
        index[key] = {
            "name": name,
            "crate_id": 0,
            "visibility": visibility,
            "inner": {kind: inner},
        }
        if span:
            index[key]["span"] = {}
        paths[key] = {"path": list(path), "kind": kind, "crate_id": 0}

    add(0, "fixture", ("fixture",), "module", {"items": [1, 4, 5, 6], "is_crate": True})
    add(1, "OddOwner", ("fixture", "OddOwner"), "struct",
        {"fields": [], "impls": [2]})
    add(2, "OddOwner", ("fixture", "OddOwner"), "impl",
        {"items": [3]}, span=True)
    add(3, "apply_change", ("fixture", "OddOwner", "apply_change"), "function", {
        "sig": {
            "inputs": [["self", {"borrowed_ref": {"is_mutable": True}}]],
            "output": {"primitive": "unit"},
        },
        "generics": {"params": []},
    })
    add(4, "AliasOwner", ("fixture", "AliasOwner"), "use", {
        "name": "AliasOwner", "id": 1, "is_glob": False,
    })
    add(5, "OwnerAlias", ("fixture", "OwnerAlias"), "type_alias", {
        "type": {"resolved_path": {"id": 1, "args": None}},
    })

    add(6, "builder", ("fixture", "builder"), "module",
        {"items": [7], "is_crate": False})
    add(7, "Worksheet", ("fixture", "builder", "Worksheet"), "struct",
        {"fields": [], "impls": [8]})
    add(8, "Worksheet", ("fixture", "builder", "Worksheet"), "impl",
        {"items": [9]}, span=True)
    add(9, "apply_change", ("fixture", "builder", "Worksheet", "apply_change"), "function", {
        "sig": {
            "inputs": [["self", {"borrowed_ref": {"is_mutable": True}}]],
            "output": {"primitive": "unit"},
        },
        "generics": {"params": []},
    })
    return {"root": 0, "includes_private": False, "index": index, "paths": paths}


def delegated_glob_fixture() -> dict:
    """Add a foreign glob whose target module is intentionally not embedded."""
    document = owner_fixture()
    document["index"]["0"]["inner"]["module"]["items"].append(10)
    document["index"]["10"] = {
        "name": "external",
        "crate_id": 0,
        "visibility": "public",
        "inner": {"use": {"name": "external", "id": 11, "is_glob": True}},
    }
    document["paths"]["11"] = {
        "path": ["external", "module"],
        "kind": "module",
        "crate_id": 99,
    }
    return document


class InventoryOwnerCoverageTests(unittest.TestCase):
    def test_policy_owner_name_outside_heuristic_covers_alias_once(self):
        document = owner_fixture()
        surface = inventory.Inventory(document, "fixture")
        owners = inventory.policy_owner_paths(
            surface, "fixture", {"fixture::OddOwner"})
        self.assertIn(("fixture", "OddOwner"), owners)
        self.assertIn(("fixture", "AliasOwner"), owners)
        self.assertIn(("fixture", "OwnerAlias"), owners)

        findings = list(inventory.rows(
            document, "fixture", result_rules=[],
            immutable_owners={"fixture::OddOwner"},
        ))
        owner_rows = [row for row in findings if row[0] == "mut_owner"]
        self.assertEqual(len(owner_rows), 1, findings)
        self.assertEqual(owner_rows[0][1], "fixture::AliasOwner::apply_change")

    def test_policy_owner_replay_uses_same_identity_and_draft_is_excluded(self):
        findings = list(inventory.rows(
            owner_fixture(), "fixture", result_rules=[],
            immutable_owners={"fixture::OddOwner", "fixture::builder::Worksheet"},
        ))
        self.assertIn(
            ("replay_return", "fixture::AliasOwner::apply_change",
             "return contract unverified", "register and satisfy Result<family::Commit> in result_methods"),
            findings,
        )
        self.assertNotIn(
            "fixture::builder::Worksheet::apply_change",
            {row[1] for row in findings if row[0] in {"mut_owner", "mut_draft"}},
        )

    def test_policy_owner_paths_are_read_from_profiles(self):
        profiles = [{"package": "fixture", "immutable_owners": ["fixture::OddOwner"]}]
        self.assertEqual(
            inventory.policy_immutable_owners("fixture", profiles),
            frozenset({("fixture", "OddOwner")}),
        )

    def test_policy_delegated_globs_union_all_profiles_for_crate(self):
        profiles = [
            {"package": "fixture", "delegated_globs": ["external"]},
            {"package": "fixture", "delegated_globs": ["other"]},
            {"package": "another", "delegated_globs": ["ignored"]},
        ]
        self.assertEqual(
            inventory.policy_delegated_globs("fixture", profiles),
            frozenset({"external", "other"}),
        )

    def test_missing_delegated_inputs_are_reported(self):
        profiles = [{"package": "fixture", "delegated_globs": ["external"]}]
        self.assertEqual(
            inventory.missing_delegated_inputs({"fixture"}, profiles),
            {"fixture": ("external",)},
        )
        self.assertEqual(
            inventory.missing_delegated_inputs({"fixture", "external"}, profiles),
            {},
        )

    def test_facade_policy_covers_all_all_features_foreign_glob_roots(self):
        profiles = inventory.gate.checked_policy(
            inventory.gate.read_json(inventory.gate.POLICY)
        )
        expected = frozenset({
            "litchi_cfb", "litchi_core", "litchi_crypto", "litchi_doc",
            "litchi_docx", "litchi_drawingml", "litchi_eval", "litchi_formula",
            "litchi_fonts", "litchi_imgconv", "litchi_keynote", "litchi_odf_common",
            "litchi_odp", "litchi_ods", "litchi_odraw", "litchi_odt",
            "litchi_ole_common", "litchi_opc", "litchi_pages", "litchi_ppt",
            "litchi_pptx", "litchi_rtf", "litchi_sheet", "litchi_sheet_text",
            "litchi_sign", "litchi_xls", "litchi_xlsb", "litchi_xlsx",
        })
        self.assertTrue(
            expected <= inventory.policy_delegated_globs("litchi", profiles),
            inventory.policy_delegated_globs("litchi", profiles),
        )

    def test_declared_delegated_glob_is_processed_without_hiding_undeclared_glob(self):
        document = delegated_glob_fixture()
        with self.assertRaisesRegex(ValueError, "glob target requires module"):
            list(inventory.rows(
                document, "fixture", result_rules=[], immutable_owners=set(),
                delegated_globs=frozenset(),
            ))
        findings = list(inventory.rows(
            document, "fixture", result_rules=[],
            immutable_owners={"fixture::OddOwner"},
            delegated_globs=frozenset({"external"}),
        ))
        self.assertIn(
            ("replay_return", "fixture::AliasOwner::apply_change",
             "return contract unverified",
             "register and satisfy Result<family::Commit> in result_methods"),
            findings,
        )

    def test_missing_policy_owner_fails_closed(self):
        surface = inventory.Inventory(owner_fixture(), "fixture")
        with self.assertRaisesRegex(inventory.InventoryError, "MissingOwner missing"):
            inventory.policy_owner_paths(surface, "fixture", {"fixture::MissingOwner"})
        with self.assertRaisesRegex(inventory.InventoryError, "MissingOwner missing"):
            list(inventory.rows(
                owner_fixture(), "fixture", result_rules=[],
                immutable_owners={"fixture::MissingOwner"},
            ))

    def test_ambiguous_policy_owner_fails_closed(self):
        document = owner_fixture()
        document["index"]["0"]["inner"]["module"]["items"].append(10)
        document["index"]["10"] = {
            "name": "CyclicAlias",
            "crate_id": 0,
            "visibility": "public",
            "inner": {"type_alias": {
                "type": {"resolved_path": {"id": 10, "args": None}},
            }},
        }
        document["paths"]["10"] = {
            "path": ["fixture", "CyclicAlias"],
            "kind": "type_alias",
            "crate_id": 0,
        }
        surface = inventory.Inventory(document, "fixture")
        with self.assertRaisesRegex(inventory.InventoryError, "CyclicAlias.*ambiguous"):
            inventory.policy_owner_paths(surface, "fixture", {"fixture::CyclicAlias"})


if __name__ == "__main__":
    unittest.main()
