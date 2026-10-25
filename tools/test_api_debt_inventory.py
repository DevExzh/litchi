from __future__ import annotations

import copy
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


def delegated_debt_fixture() -> dict:
    """A supplied delegated-crate document with the same owner debt."""
    document = owner_fixture()
    document["index"]["0"]["name"] = "external"
    for entry in document["paths"].values():
        entry["path"] = ["external", *entry["path"][1:]]
    return document


def duplicate_edit_fixture() -> dict:
    """Two nominal items emit one exact row through the same public path."""
    index = {
        "0": {
            "name": "fixture", "crate_id": 0, "visibility": "public",
            "inner": {"module": {"items": [1, 2], "is_crate": True}},
        },
        "1": {
            "name": "DuplicateEdit", "crate_id": 0, "visibility": "public",
            "inner": {"struct": {"fields": [], "impls": [3]}},
        },
        "2": {
            "name": "DuplicateEdit", "crate_id": 0, "visibility": "public",
            "inner": {"struct": {"fields": [], "impls": [4]}},
        },
        "3": {
            "name": "DuplicateEdit", "crate_id": 0, "visibility": "public",
            "inner": {"impl": {"items": [5]}}, "span": {},
        },
        "4": {
            "name": "DuplicateEdit", "crate_id": 0, "visibility": "public",
            "inner": {"impl": {"items": [6]}}, "span": {},
        },
        "5": {
            "name": "apply_change", "crate_id": 0, "visibility": "public",
            "inner": {"function": {
                "sig": {"inputs": [["self", {"borrowed_ref": {"is_mutable": True}}]],
                        "output": {"primitive": "unit"}},
                "generics": {"params": []},
            }},
        },
        "6": {
            "name": "apply_change", "crate_id": 0, "visibility": "public",
            "inner": {"function": {
                "sig": {"inputs": [["self", {"borrowed_ref": {"is_mutable": True}}]],
                        "output": {"primitive": "unit"}},
                "generics": {"params": []},
            }},
        },
    }
    paths = {
        "0": {"path": ["fixture"], "kind": "module", "crate_id": 0},
        "1": {"path": ["fixture", "DuplicateEdit"], "kind": "struct", "crate_id": 0},
        "2": {"path": ["fixture", "DuplicateEdit"], "kind": "struct", "crate_id": 0},
        "3": {"path": ["fixture", "DuplicateEdit"], "kind": "impl", "crate_id": 0},
        "4": {"path": ["fixture", "DuplicateEdit"], "kind": "impl", "crate_id": 0},
        "5": {"path": ["fixture", "DuplicateEdit", "apply_change"],
              "kind": "function", "crate_id": 0},
        "6": {"path": ["fixture", "DuplicateEdit", "apply_change"],
              "kind": "function", "crate_id": 0},
    }
    return {"root": 0, "includes_private": False, "index": index, "paths": paths}


def commit_shape_fixture(crate: str, modules: tuple[str, ...], methods: tuple[str, ...]) -> dict:
    """Build one small commit family without borrowing a format graph."""
    index = {}
    paths = {}

    def add(item_id, name, path, kind, inner, *, visibility="public", span=False):
        key = str(item_id)
        index[key] = {
            "name": name, "crate_id": 0, "visibility": visibility,
            "inner": {kind: inner},
        }
        if span:
            index[key]["span"] = {}
        paths[key] = {"path": list(path), "kind": kind, "crate_id": 0}

    add(0, crate, (crate,), "module", {"items": [], "is_crate": True})
    parent_id = 0
    parent_path = (crate,)
    next_id = 1
    for module in modules:
        module_id = next_id
        next_id += 1
        module_path = parent_path + (module,)
        add(module_id, module, module_path, "module", {"items": [], "is_crate": False})
        index[str(parent_id)]["inner"]["module"]["items"].append(module_id)
        parent_id = module_id
        parent_path = module_path

    commit_id = next_id
    impl_id = commit_id + 1
    function_ids = list(range(impl_id + 1, impl_id + 1 + len(methods)))
    commit_path = parent_path + ("Commit",)
    add(commit_id, "Commit", commit_path, "struct", {"fields": [], "impls": [impl_id]})
    index[str(parent_id)]["inner"]["module"]["items"].append(commit_id)
    add(impl_id, "Commit", commit_path, "impl", {"items": function_ids}, span=True)
    for function_id, method in zip(function_ids, methods):
        add(function_id, method, commit_path + (method,), "function", {
            "sig": {"inputs": [], "output": {"primitive": "unit"}},
            "generics": {"params": []},
        })
    return {"root": 0, "includes_private": False, "index": index, "paths": paths}


class InventoryOwnerCoverageTests(unittest.TestCase):
    def test_nominal_coverage_requires_the_local_definition_identity(self):
        methods = ("package", "into_package", "patch", "diagnostics", "into_parts")
        path = ("fixture", "foreign", "Commit")
        valid = commit_shape_fixture("fixture", ("foreign",), methods)
        self.assertIn(path, inventory.auditable_commit_nominals(valid, "fixture"))

        foreign_index_only = copy.deepcopy(valid)
        foreign_index_only["index"]["2"]["crate_id"] = 99
        self.assertNotIn(
            path, inventory.auditable_commit_nominals(foreign_index_only, "fixture")
        )

        foreign_index_item = copy.deepcopy(valid)
        foreign_index_item["index"]["2"]["crate_id"] = 99
        foreign_index_item["paths"]["2"]["crate_id"] = 99
        self.assertNotIn(
            path, inventory.auditable_commit_nominals(foreign_index_item, "fixture")
        )

        reexport = copy.deepcopy(valid)
        reexport["index"]["2"] = {
            "name": "Commit", "crate_id": 0, "visibility": "public",
            "inner": {"type_alias": {"type": {
                "resolved_path": {"id": "foreign-target", "args": None},
            }, "generics": {"params": [], "where_predicates": []}}},
        }
        reexport["paths"]["2"]["kind"] = "type_alias"
        reexport["paths"]["foreign-target"] = {
            "crate_id": 99, "path": ["fixture", "foreign", "Commit"],
            "kind": "struct",
        }
        self.assertNotIn(
            path, inventory.auditable_commit_nominals(reexport, "fixture")
        )

    def test_file_accessors_are_scoped_to_the_unified_facade(self):
        methods = (
            "file", "into_file", "patch", "diagnostics", "value", "into_value", "into_parts",
        )
        format_document = commit_shape_fixture("fixture", ("format",), methods)
        findings = list(inventory.rows(
            format_document, "fixture", result_rules=[], immutable_owners=set(),
            delegated_globs=set(),
        ))
        commit_rows = [row for row in findings if row[0] == "commit_type"]
        self.assertEqual(len(commit_rows), 1, findings)
        self.assertIn("add into_package,package", commit_rows[0][3])
        self.assertFalse(
            inventory.gate.ReturnTypes(inventory.Inventory(format_document, "fixture")).is_commit(
                ("fixture", "format", "Commit")
            )
        )

        facade_document = commit_shape_fixture("litchi_ops", ("transaction",), methods)
        facade_findings = list(inventory.rows(
            facade_document, "litchi_ops", result_rules=[], immutable_owners=set(),
            delegated_globs=set(),
        ))
        self.assertNotIn("commit_type", {row[0] for row in facade_findings}, facade_findings)
        self.assertTrue(
            inventory.gate.ReturnTypes(inventory.Inventory(facade_document, "litchi_ops")).is_commit(
                ("litchi_ops", "transaction", "Commit")
            )
        )

    def test_facade_commit_requires_both_value_accessors_and_rejects_legacy_results(self):
        required = (
            "file", "into_file", "patch", "diagnostics", "value", "into_value", "into_parts",
        )
        path = ("litchi_ops", "transaction", "Commit")
        for missing in required:
            with self.subTest(missing=missing):
                methods = tuple(method for method in required if method != missing)
                document = commit_shape_fixture("litchi_ops", ("transaction",), methods)
                return_types = inventory.gate.ReturnTypes(
                    inventory.Inventory(document, "litchi_ops")
                )
                self.assertFalse(return_types.is_commit(path))
                rows = list(inventory.rows(
                    document, "litchi_ops", result_rules=[], immutable_owners=set(),
                    delegated_globs=set(),
                ))
                commit_rows = [row for row in rows if row[0] == "commit_type"]
                self.assertEqual(len(commit_rows), 1, rows)
                self.assertIn("add " + missing, commit_rows[0][3])

        legacy = (
            "file", "into_file", "patch", "diagnostics", "results", "into_results", "into_parts",
        )
        document = commit_shape_fixture("litchi_ops", ("transaction",), legacy)
        return_types = inventory.gate.ReturnTypes(
            inventory.Inventory(document, "litchi_ops")
        )
        self.assertFalse(return_types.is_commit(path))
        rows = list(inventory.rows(
            document, "litchi_ops", result_rules=[], immutable_owners=set(),
            delegated_globs=set(),
        ))
        commit_rows = [row for row in rows if row[0] == "commit_type"]
        self.assertEqual(len(commit_rows), 1, rows)
        self.assertIn("add into_value,value", commit_rows[0][3])

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
            ("mut_owner", "fixture::AliasOwner::apply_change", "&mut self",
             "change::Edit via Package::edit_change() / apply_change(&Patch)"),
            findings,
        )
        self.assertIn(
            ("replay_return", "fixture::AliasOwner::apply_change",
             "return contract unverified",
             "register and satisfy Result<family::Commit> in result_methods"),
            findings,
        )

    def test_supplied_delegated_document_is_checked_for_debt(self):
        findings = list(inventory.rows(
            delegated_debt_fixture(), "external", result_rules=[],
            immutable_owners={"external::OddOwner"}, delegated_globs=set(),
        ))
        self.assertTrue(
            any(row[0] == "mut_owner" and
                row[1] == "external::AliasOwner::apply_change"
                for row in findings),
            findings,
        )

    def test_exact_duplicate_rows_are_emitted_once(self):
        findings = list(inventory.rows(
            duplicate_edit_fixture(), "fixture", result_rules=[],
            immutable_owners=set(), delegated_globs=set(),
        ))
        expected = (
            "edit_type", "fixture::DuplicateEdit", "",
            "rename to the focused module's Edit",
        )
        self.assertEqual(findings, [expected])

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
