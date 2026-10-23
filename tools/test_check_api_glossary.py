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
from tools import api_debt_inventory as inventory


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
pub struct Owner;
impl Owner {
    pub fn open() -> Self { Owner }
    pub fn save(&self) {}
    pub fn put_part(&mut self) {}
    pub fn into_bytes(self) {}
    fn private_mut(&mut self) {}
}
pub struct Frozen;
impl Frozen { pub fn save(&self) {} pub fn into_bytes(self) {} }
pub mod strict {
    pub struct Package;
    pub struct Edit;
    impl Edit { pub fn commit(self) -> Result<Commit, Error> { Ok(Commit) } pub fn set(&mut self) {} }
    pub struct Patch;
    impl Patch { pub fn is_noop(&self) -> bool { true } pub fn inverse(&self) -> Self { Patch } }
    pub struct Commit;
    impl Commit {
        pub fn package(&self) -> &Package { &Package }
        pub fn into_package(self) -> Package { Package }
        pub fn patch(&self) -> &Patch { &Patch }
        pub fn diagnostics(&self) -> &Diagnostics { &Diagnostics }
        pub fn into_parts(self) -> (Package, Patch) { (Package, Patch) }
    }
    pub struct Diagnostics;
    impl Diagnostics { pub fn changed(&self) -> bool { false } }
    pub enum Error { Bad }
    pub enum LimitKind { Bytes }
}
pub mod loose {
    pub struct Edit;
    impl Edit { pub fn commit(self) {} pub fn rollback(self) {} }
    pub struct Patch;
    pub struct Commit;
    impl Commit { pub fn snapshot(&self) {} }
    pub struct Diagnostics;
    pub enum Error { Bad }
    pub enum LimitKind { Bytes }
}
pub mod shared {
    pub use crate::strict::{Patch, Error, LimitKind};
    pub struct Edit;
    impl Edit { pub fn commit(self) -> Result<Commit, Error> { Ok(Commit) } }
    pub use crate::strict::Commit;
    pub use std::time::Instant as Diagnostics;
}
pub mod unrelated_commit {
    pub use crate::strict::{Patch, Diagnostics, Error, LimitKind};
    pub type Commit = std::time::Duration;
    pub struct Edit;
    impl Edit { pub fn commit(self) -> Result<Commit, Error> { Ok(Commit::ZERO) } }
}
pub mod staged_output {
    pub use crate::strict::{Package, Patch, Commit, Diagnostics, Error, LimitKind};
    pub struct Edit;
    impl Edit {
        pub fn commit(self) -> Result<Commit, Error> { Ok(Commit) }
        pub fn put_stream_shared_with_rendered(&mut self) -> Vec<u8> { vec![] }
        pub fn put_streams_shared_with_rendered(&mut self) -> Option<Vec<u8>> { None }
        pub fn write_to(&self) {}
        pub fn to_bytes(&self) -> Vec<u8> { vec![] }
        pub fn save(&self) {}
        pub fn publish_to_stream(self) {}
        pub fn prepare(&self) {}
        fn to_plain_bytes(&self) -> Vec<u8> { vec![] }
    }
}
pub use staged_output::Edit as StagedEdit;
pub mod builder {
    pub struct Edit;
    impl Edit { pub fn to_bytes(&self) -> Vec<u8> { vec![] } }
}
impl strict::Package { pub fn to_bytes(&self) -> Vec<u8> { vec![] } }
impl strict::Patch { pub fn to_bytes(&self) -> Vec<u8> { vec![] } }
pub mod closed_error { pub enum Error { Invalid } }
pub mod extensible_error { #[non_exhaustive] pub enum Error { Invalid } }
pub fn with_limits() {}
pub fn open_with_limits() {}
pub fn open_with_options() {}
pub struct MutableDraft;
pub mod returns {
    pub struct Package;
    pub struct Other;
    pub struct Patch;
    pub struct Error;
    pub struct Triad<P>(pub P);
    pub type Commit<P = Package> = Triad<P>;
    pub type Result<T, E = Error> = std::result::Result<T, E>;
    pub type Swapped<E, T> = Result<T, E>;
    pub type OwnerAlias = Package;
    pub use Package as OwnerExport;
    pub mod fake { pub struct Result<T>(pub T); pub struct Commit; }
    impl Package {
        pub fn apply(&self, _: &Patch) -> Result<Commit> { Ok(Triad(Package)) }
        pub fn apply_bad(&self, _: &Patch) -> Result<Self> { Ok(Package) }
        pub fn apply_other_patch(&self, _: &Other) -> Result<Commit> { Ok(Triad(Package)) }
        pub fn apply_other_owner(&self, _: &Patch) -> Result<Commit<Other>> { Ok(Triad(Other)) }
        pub fn good(&self) -> Swapped<Error, Commit> { Ok(Triad(Package)) }
        pub fn bad_owner(&self) -> Result<Self> { Ok(Package) }
        pub fn wrong_generic(&self) -> Result<Commit<Other>> { Ok(Triad(Other)) }
        pub fn direct(&self) -> Commit { Triad(Package) }
        pub fn optional(&self) -> Option<Result<Commit>> { None }
        pub fn borrowed(&self) -> Result<&Commit> { todo!() }
        pub fn fake_result(&self) -> fake::Result<Commit> { fake::Result(Triad(Package)) }
        pub fn fake_commit(&self) -> Result<fake::Commit> { Ok(fake::Commit) }
    }
    impl<P> Triad<P> {
        pub fn again(self) -> Result<Self> { Ok(self) }
        pub fn package(&self) -> &P { &self.0 }
        pub fn into_package(self) -> P { self.0 }
        pub fn patch(&self) -> &Patch { &Patch }
        pub fn diagnostics(&self) -> &() { &() }
        pub fn into_parts(self) -> (P, Patch) { (self.0, Patch) }
    }
}
pub mod durable {
    pub struct Package;
    pub struct Other;
    pub struct Error;
    impl Package {
        pub fn apply<Mode>(&self) -> Result<litchi_core::edit::Commit<Self, litchi_core::patch::Patch<Mode>>, Error> { todo!() }
        pub fn apply_other<Mode>(&self) -> Result<litchi_core::edit::Commit<Other, litchi_core::patch::Patch<Mode>>, Error> { todo!() }
        pub fn apply_fixed<Mode>(&self) -> Result<litchi_core::edit::Commit<Self, litchi_core::patch::Patch<Other>>, Error> { todo!() }
        pub fn apply_swapped<Mode>(&self) -> Result<litchi_core::edit::Commit<litchi_core::patch::Patch<Mode>, Self>, Error> { todo!() }
        pub fn apply_foreign<Mode>(&self) -> Result<litchi_core::unrelated::Commit<Self, litchi_core::patch::Patch<Mode>>, Error> { todo!() }
    }
}
pub mod shared_core {
    pub type Commit = litchi_core::edit::Commit<crate::strict::Package, crate::strict::Patch, crate::strict::Diagnostics>;
    pub use litchi_core::edit::Diagnostics;
    pub use litchi_core::edit::Commit as ReexportCommit;
    pub type RichCommit = litchi_core::edit::Commit<crate::strict::Package, crate::strict::Patch, crate::strict::Diagnostics>;
    pub use litchi_core::unrelated::Commit as UnrelatedCommit;
    pub use std::time::Instant as FakeDiagnostics;
}
'''

DEPENDENCY_SOURCE = r'''
pub mod edit {
    pub struct Diagnostics;
    pub struct Commit<P, T, D = Diagnostics>(pub P, pub T, pub D);
}
pub mod patch { pub struct Patch<Mode>(pub Mode); }
pub mod unrelated { pub struct Commit<P, T>(pub P, pub T); }
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
            {"version": 1, "profiles": [profile(banned_names=[{**ban("Id"), "match": "infix"}])]},
            {"version": 1, "profiles": [profile(family_methods="yes")]},
            {"version": 1, "profiles": [profile(immutable_owners="fixture::Owner")]},
            {"version": 1, "profiles": [profile(immutable_owners=["other::Owner"])]},
            {"version": 1, "profiles": [profile(immutable_owners=["fixture::Owner", "fixture::Owner"])]},
            {"version": 1, "profiles": [profile(result_methods="yes")]},
            {"version": 1, "profiles": [profile(result_methods=[{"method": "fixture::f"}])]},
            {"version": 1, "profiles": [profile(result_methods=[{"method": "fixture::f", "ok": "Commit"}])]},
            {"version": 1, "profiles": [profile(result_methods=[{"method": "fixture::f", "ok": "fixture::Package"}])]},
            {"version": 1, "profiles": [profile(result_methods=[{"method": "fixture::f", "ok": "fixture::Commit", "type_args": "P"}])]},
            {"version": 1, "profiles": [profile(result_methods=[{"method": "fixture::f", "ok": "fixture::Commit", "type_args": ["Package"]}])]},
            {"version": 1, "profiles": [profile(result_methods=[{"method": "fixture::f", "ok": "fixture::Commit"}] * 2)]},
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

    def test_structured_type_rules_require_exact_keys_and_qualified_paths(self):
        def policy(arg):
            return {"version": 1, "profiles": [profile(result_methods=[{
                "method": "fixture::durable::Package::apply", "ok": "litchi_core::edit::Commit",
                "type_args": ["fixture::durable::Package", arg],
            }])]}
        self.assertTrue(gate.checked_policy(policy({"path": "litchi_core::patch::Patch", "type_args": [{"generic": "Mode"}]})))
        for bad in ({}, {"generic": ""}, {"generic": "Mode", "path": "fixture::Patch"},
                    {"generic": "Self"}, {"generic": "Mode::T"}, {"path": "Patch"},
                    {"path": "fixture::Patch", "type_args": "Mode"},
                    {"path": "fixture::Patch", "type_args": [None]},
                    {"path": "fixture::Patch", "typo": []}):
            with self.subTest(bad=bad), self.assertRaises(ValueError):
                gate.checked_policy(policy(bad))

    def test_shared_type_policy_is_explicit_and_fail_closed(self):
        def policy(rows):
            return {"version": 1, "profiles": [profile(shared_types=rows)]}

        valid = [{
            "path": "fixture::shared_core::Commit",
            "target": "litchi_core::edit::Commit",
        }]
        self.assertTrue(gate.checked_policy(policy(valid)))
        for bad in (
            "not-an-array",
            [{"path": "fixture::shared_core::Commit"}],
            [{"path": "other::Commit", "target": "litchi_core::edit::Commit"}],
            [{"path": "fixture::shared_core::Commit", "target": "other::Commit"}],
            [{"path": "fixture::shared_core::Commit", "target": "litchi_core::edit::Commit"},
             {"path": "fixture::shared_core::Commit", "target": "litchi_core::edit::Diagnostics"}],
        ):
            with self.subTest(bad=bad), self.assertRaises(ValueError):
                gate.checked_policy(policy(bad))


class RustdocTests(unittest.TestCase):
    """Use real rustdoc graphs so re-export/visibility assumptions are tested."""

    @classmethod
    def setUpClass(cls):
        cls.scratch = tempfile.TemporaryDirectory(prefix="litchi-glossary-test-", dir=gate.ROOT)
        cls.addClassCleanup(cls.scratch.cleanup)
        cls.directory = Path(cls.scratch.name)
        dependency = cls.directory / "dependency.rs"
        dependency.write_text(DEPENDENCY_SOURCE, encoding="utf-8")
        library = cls.directory / "liblitchi_core.rlib"
        subprocess.run(["rustc", "--edition=2024", "--crate-type=rlib", "--crate-name=litchi_core",
                        str(dependency), "-o", str(library)], check=True, capture_output=True, text=True)
        source = cls.directory / "lib.rs"
        source.write_text(SOURCE, encoding="utf-8")
        subprocess.run(
            ["rustdoc", "--edition=2024", "--crate-type=lib", "--crate-name=fixture",
             "-Zunstable-options", "--output-format=json", "--document-hidden-items",
             "--extern", f"litchi_core={library}", "-o", str(cls.directory), str(source)],
            env={**os.environ, "RUSTC_BOOTSTRAP": "1"}, check=True,
            capture_output=True, text=True,
        )
        cls.json_path = cls.directory / "fixture.json"
        cls.document = gate.read_json(cls.json_path)

    def check(self, *names, scope="fixture", families=()):
        return gate.violations(self.document, profile(
            banned_names=[ban(n, scope) for n in names], families=list(families)))

    def test_inventory_catches_exact_and_suffixed_limit_twins(self):
        limits = {path for kind, path, _, _ in inventory.rows(self.document, "fixture")
                  if kind == "with_limits"}
        self.assertEqual(limits, {"fixture::with_limits", "fixture::open_with_limits"})
        self.assertTrue(self.check("with_limits"))

    def test_inventory_reports_staged_output_through_reexport(self):
        outputs = {path for kind, path, _, _ in inventory.rows(self.document, "fixture")
                   if kind == "edit_output"}
        self.assertEqual(outputs, {
            "fixture::StagedEdit::" + name for name in (
                "put_stream_shared_with_rendered", "put_streams_shared_with_rendered",
                "write_to", "to_bytes", "save", "publish_to_stream",
            )
        })

    def test_inventory_requires_non_exhaustive_error_enums(self):
        errors = {path for kind, path, _, _ in inventory.rows(self.document, "fixture")
                  if kind == "exhaustive_error"}
        self.assertIn("fixture::closed_error::Error", errors)
        self.assertNotIn("fixture::extensible_error::Error", errors)

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

    def test_registered_replay_checks_published_owner(self):
        failures = gate.violations(self.document, profile(result_methods=[{
            "method": "fixture::returns::Package::apply_other_owner",
            "ok": "fixture::returns::Commit", "type_args": ["fixture::returns::Other"],
        }]))
        self.assertEqual(len(failures), 1, failures)

    def test_registered_replay_supports_distinct_durable_input_patch(self):
        self.assertEqual(gate.violations(self.document, profile(result_methods=[{
            "method": "fixture::returns::Package::apply_other_patch",
            "ok": "fixture::returns::Commit",
        }])), [])

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

    def test_a_delegated_external_glob_is_left_to_its_crate_profile(self):
        # Record 0306: the facade re-exports leaf crates by glob; a profile
        # names the crates whose own profiles check those names.
        doc = copy.deepcopy(self.document)
        use = next(v["inner"]["use"] for v in doc["index"].values()
                   if v.get("inner", {}).get("use", {}).get("is_glob"))
        use["id"] = "external-module"
        doc["paths"]["external-module"] = {
            "crate_id": 99, "kind": "module", "path": ["external", "module"],
        }
        self.assertEqual(gate.violations(doc, profile(delegated_globs=["external"])), [])
        with self.assertRaisesRegex(ValueError, "glob target requires module"):
            gate.violations(doc, profile(delegated_globs=["other"]))

    def test_delegated_globs_must_be_crate_names(self):
        policy = {"version": 1, "profiles": [profile(delegated_globs=["not a crate"])]}
        with self.assertRaisesRegex(ValueError, "delegated_globs"):
            gate.checked_policy(policy)

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

    def test_suffix_and_prefix_bans_match_whole_name_ends(self):
        suffix = {**ban("_with_limits"), "match": "suffix"}
        failures = gate.violations(self.document, profile(banned_names=[suffix]))
        self.assertEqual(len(failures), 1, failures)
        self.assertIn("fixture::open_with_limits", failures[0])
        prefix = {**ban("Mutable"), "match": "prefix"}
        failures = gate.violations(self.document, profile(banned_names=[prefix]))
        self.assertEqual(len(failures), 1, failures)
        self.assertIn("fixture::MutableDraft", failures[0])
        # A prefix or suffix ban never matches the bare affix itself.
        exact_affix = {**ban("open"), "match": "prefix"}
        self.assertEqual(len(gate.violations(self.document, profile(banned_names=[exact_affix]))), 2)
        self.assertEqual(self.check("_with_limits"), [])
        # The same name may be banned exactly and as an affix in one scope.
        policy = {"version": 1, "profiles": [profile(banned_names=[ban("Mutable"), prefix])]}
        self.assertTrue(gate.checked_policy(policy))

    def test_immutable_owner_rejects_public_mut_self_methods_only(self):
        failures = gate.violations(self.document, profile(immutable_owners=["fixture::Owner"]))
        self.assertEqual(len(failures), 1, failures)
        self.assertIn("fixture::Owner::put_part takes &mut self", failures[0])
        self.assertEqual(gate.violations(self.document, profile(immutable_owners=["fixture::Frozen"])), [])
        failures = gate.violations(self.document, profile(immutable_owners=["fixture::Missing"]))
        self.assertEqual(failures, ["missing immutable owner type fixture::Missing"])

    def test_family_methods_require_the_commit_triad_accessors(self):
        families = ["fixture::strict", "fixture::loose"]
        self.assertEqual(gate.violations(self.document, profile(families=families[:1], family_methods=True)), [])
        # Without family_methods only the six type names are checked.
        self.assertEqual(gate.violations(self.document, profile(families=families)), [])
        failures = gate.violations(self.document, profile(families=families, family_methods=True))
        expected = {
            "fixture::loose::Commit exposes forbidden method snapshot (record 0351)",
            "fixture::loose::Commit requires method diagnostics (record 0351)",
            "fixture::loose::Commit requires method into_package (record 0351)",
            "fixture::loose::Commit requires method into_parts (record 0351)",
            "fixture::loose::Commit requires method package (record 0351)",
            "fixture::loose::Commit requires method patch (record 0351)",
            "fixture::loose::Diagnostics requires method changed (record 0351)",
            "fixture::loose::Edit exposes forbidden method rollback (record 0351)",
            "fixture::loose::Edit::commit must return Result<fixture::loose::Commit> (record 0351)",
            "fixture::loose::Patch requires method inverse (record 0351)",
            "fixture::loose::Patch requires method is_noop (record 0351)",
        }
        self.assertEqual(set(failures), expected)

    def test_family_methods_leave_dependency_members_to_their_crate(self):
        # Re-exported dependency types carry no rustdoc methods here; their
        # defining crate checks their methods. Local family members and the
        # exact Edit::commit return type are still checked.
        self.assertEqual(
            gate.violations(self.document, profile(families=["fixture::shared"], family_methods=True)), [])

    def test_family_return_rejects_an_unrelated_dependency_named_commit(self):
        failures = gate.violations(self.document, profile(
            families=["fixture::unrelated_commit"], family_methods=True))
        self.assertEqual(failures, [
            "fixture::unrelated_commit::Edit::commit must return Result<fixture::unrelated_commit::Commit> (record 0351)",
        ])

    def test_result_shape_uses_alias_substitutions_defaults_and_nominal_identity(self):
        def rule(method, **extra):
            return {"method": "fixture::returns::" + method,
                    "ok": "fixture::returns::Commit", **extra}
        good = [rule("Package::good"), rule("OwnerAlias::good"), rule("OwnerExport::good"),
                rule("Commit::again"), rule("Package::good", type_args=["fixture::returns::Package"])]
        for case in good:
            with self.subTest(case=case):
                self.assertEqual(gate.violations(self.document, profile(result_methods=[case])), [])
        for method in ("bad_owner", "wrong_generic", "direct", "optional", "borrowed", "fake_result", "fake_commit", "missing"):
            with self.subTest(method=method):
                failures = gate.violations(self.document, profile(result_methods=[rule("Package::" + method)]))
                self.assertEqual(len(failures), 1, failures)
                self.assertIn("must return Result<", failures[0])
        failures = gate.violations(self.document, profile(result_methods=[{
            "method": "fixture::returns::Package::fake_commit", "ok": "fixture::returns::fake::Commit",
        }]))
        self.assertEqual(len(failures), 1, failures)

    def test_direct_shared_commit_preserves_nested_method_generic_identity(self):
        rule = {"method": "fixture::durable::Package::apply", "ok": "litchi_core::edit::Commit",
                "type_args": ["fixture::durable::Package", {
                    "path": "litchi_core::patch::Patch", "type_args": [{"generic": "Mode"}],
                }]}
        self.assertEqual(gate.violations(self.document, profile(result_methods=[rule])), [])
        findings = list(inventory.rows(self.document, "fixture", result_rules=[rule]))
        self.assertNotIn(rule["method"], {path for category, path, _, _ in findings if category == "replay_return"})
        for method in ("apply_other", "apply_fixed", "apply_swapped", "apply_foreign"):
            changed = {**rule, "method": "fixture::durable::Package::" + method}
            with self.subTest(method=method):
                self.assertEqual(len(gate.violations(self.document, profile(result_methods=[changed]))), 1)
        wrong_owner = {**rule, "method": "fixture::durable::Package::apply_other",
                       "type_args": ["fixture::durable::Other", rule["type_args"][1]]}
        self.assertEqual(len(gate.violations(self.document, profile(result_methods=[wrong_owner]))), 1)
        for name in ("Unknown", "Self"):
            changed = copy.deepcopy(rule)
            changed["type_args"][1]["type_args"] = [{"generic": name}]
            self.assertEqual(len(gate.violations(self.document, profile(result_methods=[changed]))), 1)
        for path in ("litchi_core::missing::Patch", "litchi_core::unrelated::Commit"):
            changed = {**rule, "ok": path}
            self.assertEqual(len(gate.violations(self.document, profile(result_methods=[changed]))), 1)
        ambiguous = copy.deepcopy(self.document)
        canonical = next(value for value in ambiguous["paths"].values()
                         if value["path"] == ["litchi_core", "edit", "Commit"])
        ambiguous["paths"]["ambiguous-core-commit"] = copy.deepcopy(canonical)
        self.assertEqual(len(gate.violations(ambiguous, profile(result_methods=[rule]))), 1)

    def test_shared_type_rows_use_nominal_ids_for_aliases_and_reexports(self):
        rows = [
            {"path": "fixture::shared_core::Commit", "target": "litchi_core::edit::Commit"},
            {"path": "fixture::shared_core::Diagnostics", "target": "litchi_core::edit::Diagnostics"},
            {"path": "fixture::shared_core::ReexportCommit", "target": "litchi_core::edit::Commit"},
            {"path": "fixture::shared_core::RichCommit", "target": "litchi_core::edit::Commit"},
        ]
        self.assertEqual(gate.violations(self.document, profile(shared_types=rows)), [])

        fake = dict(rows[0], path="fixture::shared_core::UnrelatedCommit")
        failures = gate.violations(self.document, profile(shared_types=[fake]))
        self.assertEqual(failures, [
            "fixture::shared_core::UnrelatedCommit must resolve nominally to "
            "litchi_core::edit::Commit (record 0351)",
        ])

        fake_diagnostics = dict(rows[1], path="fixture::shared_core::FakeDiagnostics")
        failures = gate.violations(self.document, profile(shared_types=[fake_diagnostics]))
        self.assertEqual(failures, [
            "fixture::shared_core::FakeDiagnostics must resolve nominally to "
            "litchi_core::edit::Diagnostics (record 0351)",
        ])

        # A format-owned diagnostics type remains valid when it is not
        # nominated as the shared core Diagnostics type.
        self.assertEqual(gate.violations(self.document, profile(
            families=["fixture::strict"], family_methods=True)), [])

    def test_shared_type_rows_reject_ambiguous_canonical_ids(self):
        document = copy.deepcopy(self.document)
        canonical = next(
            value for value in document["paths"].values()
            if value["path"] == ["litchi_core", "edit", "Commit"]
        )
        document["paths"]["ambiguous-shared-core-commit"] = copy.deepcopy(canonical)
        rows = [{"path": "fixture::shared_core::Commit", "target": "litchi_core::edit::Commit"}]
        failures = gate.violations(document, profile(shared_types=rows))
        self.assertEqual(failures, [
            "fixture::shared_core::Commit must resolve nominally to "
            "litchi_core::edit::Commit (record 0351)",
        ])

    def test_unresolved_and_recursive_result_aliases_fail_closed(self):
        for mutation in ("missing", "cycle"):
            doc = copy.deepcopy(self.document)
            item_id, alias = next((k, v) for k, v in doc["index"].items()
                                  if v.get("name") == "Swapped")
            alias["inner"]["type_alias"]["type"] = {"resolved_path": {
                "id": item_id if mutation == "cycle" else "unknown-return", "args": None,
            }}
            checked = profile(result_methods=[{
                "method": "fixture::returns::Package::good", "ok": "fixture::returns::Commit",
            }])
            if mutation == "missing":
                with self.assertRaisesRegex(ValueError, "missing rustdoc item"):
                    gate.violations(doc, checked)
                continue
            failures = gate.violations(doc, checked)
            self.assertEqual(len(failures), 1, failures)

    def test_inventory_reports_unverified_replay_and_wrong_commit_results(self):
        findings = list(inventory.rows(self.document, "fixture", result_rules=[]))
        commits = {path for category, path, _, _ in findings if category == "commit_return"}
        self.assertIn("fixture::loose::Edit::commit", commits)
        self.assertNotIn("fixture::strict::Edit::commit", commits)
        replay = {path: detail for category, path, detail, _ in findings if category == "replay_return"}
        self.assertEqual(replay["fixture::returns::Package::apply"], "return contract unverified")
        self.assertEqual(replay["fixture::returns::Package::apply_bad"], "returns owner")
        verified = list(inventory.rows(self.document, "fixture", result_rules=[{
            "method": "fixture::returns::Package::apply", "ok": "fixture::returns::Commit",
        }]))
        remaining = {path for category, path, _, _ in verified if category == "replay_return"}
        self.assertNotIn("fixture::returns::Package::apply", remaining)
        self.assertIn("fixture::returns::Package::apply_bad", remaining)

    def test_family_methods_reject_staged_artifact_output(self):
        failures = gate.violations(self.document, profile(
            families=["fixture::staged_output"], family_methods=True))
        self.assertEqual(set(failures), {
            f"fixture::staged_output::Edit exposes forbidden method {name} (record 0351)"
            for name in ("put_stream_shared_with_rendered", "put_streams_shared_with_rendered",
                         "write_to", "to_bytes", "save", "publish_to_stream")
        })

    def test_inventory_cli_uses_graph_identity_for_renamed_captures(self):
        capture = self.directory / "renamed-full-profile.json"
        capture.write_text(json.dumps(self.document))
        output = io.StringIO()
        with contextlib.redirect_stdout(output):
            self.assertEqual(inventory.main([str(capture)]), 0)
        self.assertIn("fixture\treplay_return\tfixture::returns::Package::apply_bad", output.getvalue())

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
