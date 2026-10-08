from __future__ import annotations

import copy
import contextlib
import io
import json
import os
from pathlib import Path
import subprocess
import tempfile
import unittest
from unittest.mock import patch

from tools import check_api_glossary as gate


SOURCE = r'''
#![allow(dead_code)]
mod hidden {
    pub struct Renamed;
    impl Renamed {
        pub fn old_method(&self) {}
        fn secret_method(&self) {}
    }
    pub struct Secret;
}
pub use hidden::Renamed as Package;
pub use hidden::Renamed as OtherPackage;
pub type Alias = hidden::Renamed;
pub mod tx {
    pub struct Edit;
    pub struct Patch;
    pub struct Commit;
    pub struct Diagnostics;
    pub enum Error { Bad }
    pub enum LimitKind { Bytes }
}
pub use tx as family;
pub mod glob { pub use crate::tx::*; }
pub mod incomplete {
    pub struct Edit;
    pub fn Patch() {}
    pub mod nested { pub struct Commit; }
}
pub enum Selector { Name, Id }
pub mod domain { pub struct Id; pub struct ReadLimits; }
pub struct ReadLimits;
pub struct Fields { pub old_field: u8, secret_field: u8 }
pub enum Variants { Named { old_variant_field: u8 } }
pub trait Contract {
    type OldAssociated;
    const OLD_CONST: u8;
    fn old_trait_method(&self);
}
#[doc(hidden)]
pub fn hidden_old() {}
pub fn leaked() -> hidden::Renamed { hidden::Renamed }
pub fn foreign() -> std::time::Duration { std::time::Duration::ZERO }
/// old_comment is documentation, not a name.
pub fn fine(old_argument: u8) { let _ = old_argument; }
pub mod cycle { pub use crate::cycle as again; }
'''


def profile(**changes):
    value = {
        "id": "fixture", "package": "fixture", "default_features": False,
        "features": [], "banned_names": [], "families": [],
    }
    value.update(changes)
    return value


def ban(name, scope="fixture"):
    return {"name": name, "scope": scope, "replacement": "use the canonical name"}


class PolicyTests(unittest.TestCase):
    def test_repository_policy_has_build_profiles(self):
        profiles = gate.checked_policy(gate.read_json(gate.POLICY))
        self.assertTrue(profiles)

    def test_rejects_policy_typos_and_invalid_scope(self):
        cases = [
            {"version": 2, "profiles": []},
            {"version": True, "profiles": []},
            {"version": 1, "profiles": [], "bans": []},
            {"version": 1, "profiles": [profile(unknown=[])]},
            {"version": 1, "profiles": [profile(banned_names=[{**ban("Id"), "recursive": "false"}])]},
            {"version": 1, "profiles": [profile(), profile()]},
            {"version": 1, "profiles": [profile(default_features="false")]},
            {"version": 1, "profiles": [profile(features=["a", "a"])]},
            {"version": 1, "profiles": [profile(banned_names=[ban("Id", "")])]},
            {"version": 1, "profiles": [profile(banned_names=[ban("Id", "other")])]},
            {"version": 1, "profiles": [profile(banned_names=[ban("Id"), ban("Id")])]},
            {"version": 1, "profiles": [profile(families=["fixture::tx", "fixture::tx"])]},
        ]
        for value in cases:
            with self.subTest(value=value), self.assertRaises(ValueError):
                gate.checked_policy(value)

    def test_build_command_has_explicit_features_and_hidden_items(self):
        command = gate.rustdoc_command(profile(features=["one", "two"]), Path("build"))
        self.assertIn("--no-default-features", command)
        self.assertIn("one,two", command)
        self.assertIn("--document-hidden-items", command)
        self.assertNotIn("--document-private-items", command)
        self.assertNotIn("--no-default-features", gate.rustdoc_command(profile(default_features=True), Path("build")))


class RustdocTests(unittest.TestCase):
    """Use real rustdoc graphs so re-export/visibility assumptions are tested."""

    @classmethod
    def setUpClass(cls):
        cls.scratch = tempfile.TemporaryDirectory(prefix="litchi-glossary-test-", dir=gate.ROOT.parent)
        cls.addClassCleanup(cls.scratch.cleanup)
        cls.directory = Path(cls.scratch.name)
        source = cls.directory / "lib.rs"
        source.write_text(SOURCE, encoding="utf-8")
        subprocess.run(
            ["rustdoc", "--edition=2024", "--crate-type=lib", "--crate-name=fixture",
             "-Zunstable-options", "--output-format=json", "--document-hidden-items",
             "-o", str(cls.directory), str(source)],
            env={**os.environ, "RUSTC_BOOTSTRAP": "1"}, check=True,
            capture_output=True, text=True,
        )
        cls.json_path = cls.directory / "fixture.json"
        cls.document = gate.read_json(cls.json_path)

    def check(self, *names, scope="fixture", families=()):
        return gate.violations(self.document, profile(
            banned_names=[ban(n, scope) for n in names], families=list(families)))

    def test_empty_rules_pass_real_graph_with_cycles(self):
        self.assertEqual(self.check(), [])

    def test_public_methods_follow_private_module_reexports_and_aliases(self):
        failures = self.check("old_method")
        self.assertTrue(any("fixture::Package::old_method" in f for f in failures), failures)
        self.assertTrue(any("fixture::OtherPackage::old_method" in f for f in failures), failures)
        self.assertTrue(self.check("old_method", scope="fixture::Alias"))

    def test_banned_alias_name_and_nominal_type_are_both_checked(self):
        self.assertTrue(self.check("Package"))
        self.assertTrue(self.check("Renamed"))

    def test_type_reference_is_checked_without_exporting_private_siblings(self):
        self.assertTrue(self.check("Renamed", scope="fixture::hidden"))
        self.assertEqual(self.check("Secret", "secret_method", "secret_field"), [])

    def test_dependency_type_reference_is_attributed_to_exporter_scope(self):
        failures = self.check("Duration", scope="fixture::foreign")
        self.assertEqual(len(failures), 1, failures)
        self.assertIn("fixture::foreign exposes banned name Duration", failures[0])

    def test_scope_is_segment_based_and_preserves_domain_id(self):
        failures = self.check("Id", scope="fixture::Selector")
        self.assertEqual(len(failures), 1, failures)
        self.assertIn("fixture::Selector::Id", failures[0])
        self.assertIn("missing API scope fixture::Select", self.check("Id", scope="fixture::Select")[0])

    def test_nonrecursive_ban_preserves_nested_domain_names(self):
        rule = {**ban("Id"), "recursive": False}
        self.assertEqual(gate.violations(self.document, profile(banned_names=[rule])), [])
        rule["name"] = "Package"
        self.assertTrue(gate.violations(self.document, profile(banned_names=[rule])))
        rule = {**ban("Id", "fixture::Selector"), "recursive": False}
        self.assertTrue(gate.violations(self.document, profile(banned_names=[rule])))

    def test_nonrecursive_limit_ban_rejects_only_root_name(self):
        rule = {**ban("ReadLimits"), "recursive": False}
        failures = gate.violations(self.document, profile(banned_names=[rule]))
        self.assertEqual(len(failures), 1)
        self.assertIn("fixture::ReadLimits exposes", failures[0])
        self.assertEqual(len(self.check("ReadLimits")), 2)

    def test_fields_variants_trait_members_and_hidden_public_items(self):
        for name in ("old_field", "old_variant_field", "OldAssociated", "OLD_CONST", "old_trait_method", "hidden_old"):
            with self.subTest(name=name):
                self.assertTrue(self.check(name))

    def test_docs_arguments_and_substrings_are_not_public_item_names(self):
        self.assertEqual(self.check("old_comment", "old_argument", "old"), [])

    def test_six_type_family_supports_module_aliases_and_globs(self):
        self.assertEqual(self.check(families=("fixture::tx", "fixture::family", "fixture::glob")), [])

    def test_family_missing_or_wrong_kind_does_not_pass(self):
        failures = self.check(families=("fixture::missing", "fixture::incomplete"))
        self.assertEqual(len(failures), 6, failures)
        self.assertTrue(any("requires exported type Patch" in f for f in failures))
        self.assertTrue(any("requires exported type Commit" in f for f in failures))

    def test_synthetic_and_dependency_blanket_methods_are_not_checked(self):
        self.assertEqual(self.check("borrow", "try_into", "from"), [])

    def test_unresolved_external_glob_refuses_instead_of_skipping(self):
        doc = copy.deepcopy(self.document)
        use = next(v["inner"]["use"] for v in doc["index"].values()
                   if v.get("inner", {}).get("use", {}).get("is_glob"))
        use["id"] = "external-module"
        doc["paths"]["external-module"] = {
            "crate_id": 99, "kind": "module", "path": ["external", "module"],
        }
        with self.assertRaisesRegex(ValueError, "glob target requires module"):
            gate.violations(doc, profile())

    def test_rustdoc_opaque_signature_ids_do_not_imply_missing_local_types(self):
        doc = copy.deepcopy(self.document)
        function = next(v for v in doc["index"].values() if v.get("name") == "foreign")
        function["inner"]["function"]["sig"]["output"] = {
            "resolved_path": {"id": "opaque-generated-type", "path": "Opaque", "args": None},
        }
        self.assertEqual(gate.violations(doc, profile()), [])
        self.assertTrue(gate.violations(doc, profile(banned_names=[ban("Opaque", "fixture::foreign")])))

    def test_invalid_root_private_document_and_missing_child_refuse(self):
        for mutation in (
            "root", "private", "child", "root_path", "crate_id", "inner",
            "local_target", "type_payload",
        ):
            doc = copy.deepcopy(self.document)
            if mutation == "root":
                doc["root"] = "missing"
            elif mutation == "private":
                doc["includes_private"] = True
            elif mutation == "child":
                doc["index"][str(doc["root"])]["inner"]["module"]["items"].append("missing")
            elif mutation == "root_path":
                del doc["paths"][str(doc["root"])]
            elif mutation == "crate_id":
                del doc["index"][str(doc["root"])]["crate_id"]
            elif mutation == "inner":
                doc["index"][str(doc["root"])]["inner"] = None
            elif mutation == "local_target":
                target = next(k for k, v in doc["index"].items() if v["name"] == "Renamed")
                del doc["index"][target]
            elif mutation == "type_payload":
                target = next(v for v in doc["index"].values() if v["name"] == "Renamed")
                target["inner"]["struct"] = None
            with self.subTest(mutation=mutation), self.assertRaises(ValueError):
                gate.violations(doc, profile())

    def test_cli_exit_codes_and_unknown_profile(self):
        policy_path = self.directory / "policy.json"
        for bans, expected in (([], 0), ([ban("old_method")], 1)):
            policy_path.write_text(json.dumps({"version": 1, "profiles": [profile(banned_names=bans)]}))
            with contextlib.redirect_stdout(io.StringIO()), contextlib.redirect_stderr(io.StringIO()):
                self.assertEqual(gate.main(["--policy", str(policy_path), "--profile", "fixture", "--json", str(self.json_path)]), expected)
        with contextlib.redirect_stderr(io.StringIO()):
            self.assertEqual(gate.main(["--policy", str(policy_path), "--profile", "missing"]), 2)
            self.assertEqual(gate.main(["--policy", str(policy_path), "--json", str(self.json_path)]), 2)

    def test_failed_rustdoc_build_cannot_use_stale_json(self):
        policy_path = self.directory / "failure-policy.json"
        policy_path.write_text(json.dumps({"version": 1, "profiles": [profile()]}))
        with patch.object(gate.subprocess, "run", side_effect=subprocess.CalledProcessError(1, "cargo")), contextlib.redirect_stderr(io.StringIO()):
            self.assertEqual(gate.main(["--policy", str(policy_path)]), 2)


if __name__ == "__main__":
    unittest.main()
