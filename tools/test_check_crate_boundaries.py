from __future__ import annotations

import copy
import json
import tempfile
import unittest
from dataclasses import replace
from pathlib import Path

from tools import check_crate_boundaries as boundaries


def valid_snapshot(policy: boundaries.Policy) -> boundaries.Snapshot:
    edges = policy.canonical_edges | policy.migration_edges
    dependencies = {name: set() for name in policy.packages}
    normal_dependencies = {name: set() for name in policy.packages}
    for edge in edges:
        dependencies[edge.dependent].add(edge.dependency)
        if edge not in policy.dev_only_edges:
            normal_dependencies[edge.dependent].add(edge.dependency)
    dependencies["litchi-core"].update(item.name for item in policy.core_dependency_debt)
    normal_dependencies["litchi-core"].update(
        item.name for item in policy.core_dependency_debt
    )
    frozen_dependencies = {
        name: frozenset(items) for name, items in dependencies.items()
    }
    frozen_normal_dependencies = {
        name: frozenset(items) for name, items in normal_dependencies.items()
    }
    features = {name: frozenset() for name in policy.packages}
    features["litchi-core"] = frozenset(item.name for item in policy.core_feature_debt)
    facade_optional_dependencies = frozenset(
        edge.dependency
        for edge in edges
        if edge.dependent == "litchi" and edge.dependency != "litchi-core"
    )
    return boundaries.Snapshot(
        packages=policy.packages,
        manifests=frozenset(),
        edges={
            edge: (
                f"kind={'dev' if edge in policy.dev_only_edges else 'normal'}, "
                "optional=false, target=*, rename=-",
            )
            for edge in edges
        },
        dependency_kinds={
            edge: frozenset({"dev" if edge in policy.dev_only_edges else "normal"})
            for edge in edges
        },
        dependencies=frozen_dependencies,
        normal_dependencies=frozen_normal_dependencies,
        features=features,
        feature_definitions={
            "litchi": {
                "default": frozenset(),
                "all": frozenset(
                    f"dep:{dependency}" for dependency in facade_optional_dependencies
                ),
            }
        },
        normal_optional_dependencies={"litchi": facade_optional_dependencies},
    )


class BoundaryPolicyTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.raw_policy = json.loads(boundaries.DEFAULT_POLICY.read_text(encoding="utf-8"))
        cls.policy = boundaries.parse_policy(cls.raw_policy)

    def test_checked_in_policy_is_internally_consistent(self) -> None:
        self.assertEqual(boundaries.audit_snapshot(valid_snapshot(self.policy), self.policy), [])

    def test_checked_in_dev_only_edges_are_exact(self) -> None:
        self.assertEqual(
            self.policy.dev_only_edges,
            frozenset(
                {
                    boundaries.Edge("litchi-iwa-detect", "litchi-iwa-protos"),
                    boundaries.Edge("litchi-iwa-structured", "litchi-iwa-text"),
                    boundaries.Edge("litchi-sign", "soapberry-zip"),
                }
            ),
        )

    def test_dev_only_edge_cannot_be_promoted(self) -> None:
        edge = min(self.policy.dev_only_edges)
        for kind in ("build", "normal"):
            with self.subTest(kind=kind):
                snapshot = valid_snapshot(self.policy)
                evidence = dict(snapshot.edges)
                evidence[edge] = (
                    f"kind={kind}, optional=false, target=*, rename=-",
                )
                dependency_kinds = dict(snapshot.dependency_kinds)
                dependency_kinds[edge] = frozenset({kind})
                snapshot = replace(
                    snapshot,
                    edges=evidence,
                    dependency_kinds=dependency_kinds,
                )

                violations = boundaries.audit_snapshot(snapshot, self.policy)

                self.assertIn(
                    f"dev-only internal edge used outside dev: {edge.display()} "
                    f"(kind={kind}, optional=false, target=*, rename=-)",
                    violations,
                )

    def test_unannotated_dev_only_edge_is_rejected(self) -> None:
        edge = min(self.policy.dev_only_edges)
        policy = replace(
            self.policy,
            dev_only_edges=self.policy.dev_only_edges - frozenset({edge}),
        )

        violations = boundaries.audit_snapshot(valid_snapshot(self.policy), policy)

        self.assertIn(
            f"dev-only internal edge lacks policy annotation: {edge.display()}",
            violations,
        )

    def test_resolved_dev_only_edge_requires_annotation_cleanup(self) -> None:
        edge = min(self.policy.dev_only_edges)
        snapshot = valid_snapshot(self.policy)
        evidence = dict(snapshot.edges)
        del evidence[edge]
        dependency_kinds = dict(snapshot.dependency_kinds)
        del dependency_kinds[edge]
        dependencies = dict(snapshot.dependencies)
        dependencies[edge.dependent] -= frozenset({edge.dependency})
        snapshot = replace(
            snapshot,
            edges=evidence,
            dependency_kinds=dependency_kinds,
            dependencies=dependencies,
        )

        violations = boundaries.audit_snapshot(snapshot, self.policy)

        self.assertIn(
            f"resolved dev-only edge still annotated: {edge.display()}; "
            "remove its policy annotation",
            violations,
        )

    def test_dev_only_annotation_must_reference_a_canonical_edge(self) -> None:
        raw = copy.deepcopy(self.raw_policy)
        annotation = raw["dev_only_edges"][0]
        raw["packages"][annotation["dependent"]].remove(annotation["dependency"])

        with self.assertRaisesRegex(
            boundaries.PolicyError,
            "dev-only annotations must reference canonical edges",
        ):
            boundaries.parse_policy(raw)

    def test_policy_rejects_incoming_migration_host_edge(self) -> None:
        raw = copy.deepcopy(self.raw_policy)
        raw["packages"]["litchi-codepage"].append("litchi-iwa")

        with self.assertRaisesRegex(
            boundaries.PolicyError,
            "migration hosts cannot be workspace dependencies: "
            "litchi-codepage -> litchi-iwa",
        ):
            boundaries.parse_policy(raw)

    def test_snapshot_rejects_incoming_migration_host_edge(self) -> None:
        snapshot = valid_snapshot(self.policy)
        edge = boundaries.Edge("litchi-codepage", "litchi-iwa")
        evidence = dict(snapshot.edges)
        evidence[edge] = ("kind=normal, optional=false, target=*, rename=-",)
        dependency_kinds = dict(snapshot.dependency_kinds)
        dependency_kinds[edge] = frozenset({"normal"})
        dependencies = dict(snapshot.dependencies)
        dependencies[edge.dependent] |= frozenset({edge.dependency})
        normal_dependencies = dict(snapshot.normal_dependencies)
        normal_dependencies[edge.dependent] |= frozenset({edge.dependency})
        snapshot = replace(
            snapshot,
            edges=evidence,
            dependency_kinds=dependency_kinds,
            dependencies=dependencies,
            normal_dependencies=normal_dependencies,
        )

        violations = boundaries.audit_snapshot(snapshot, self.policy)

        self.assertIn(
            f"workspace dependency targets migration host: {edge.display()}",
            violations,
        )

    def test_unclassified_optional_dev_edge_is_rejected(self) -> None:
        snapshot = valid_snapshot(self.policy)
        edge = boundaries.Edge("litchi-xlsx", "litchi-doc")
        edges = dict(snapshot.edges)
        edges[edge] = ("kind=dev, optional=true, target=cfg(unix), rename=legacy",)
        dependencies = dict(snapshot.dependencies)
        dependencies[edge.dependent] |= frozenset({edge.dependency})
        snapshot = replace(snapshot, edges=edges, dependencies=dependencies)

        violations = boundaries.audit_snapshot(snapshot, self.policy)

        self.assertIn(
            "unclassified internal edge litchi-xlsx -> litchi-doc "
            "(kind=dev, optional=true, target=cfg(unix), rename=legacy)",
            violations,
        )

    def test_eval_rejects_each_normal_runtime_dependency(self) -> None:
        for runtime in sorted(self.policy.runtime_packages):
            with self.subTest(runtime=runtime):
                snapshot = valid_snapshot(self.policy)
                dependencies = dict(snapshot.dependencies)
                dependencies["litchi-eval"] |= frozenset({runtime})
                normal_dependencies = dict(snapshot.normal_dependencies)
                normal_dependencies["litchi-eval"] |= frozenset({runtime})
                snapshot = replace(
                    snapshot,
                    dependencies=dependencies,
                    normal_dependencies=normal_dependencies,
                )

                violations = boundaries.audit_snapshot(snapshot, self.policy)

                self.assertIn(
                    f"runtime-neutral crate litchi-eval depends on: {runtime}",
                    violations,
                )

    def test_eval_allows_dev_only_runtime_dependency(self) -> None:
        snapshot = valid_snapshot(self.policy)
        dependencies = dict(snapshot.dependencies)
        dependencies["litchi-eval"] |= frozenset({"tokio"})
        snapshot = replace(snapshot, dependencies=dependencies)

        violations = boundaries.audit_snapshot(snapshot, self.policy)

        self.assertEqual(violations, [])

    def test_metadata_classifies_only_normal_dependencies_as_runtime_edges(self) -> None:
        snapshot = boundaries.snapshot_from_metadata(
            {
                "workspace_members": ["litchi-eval-id"],
                "packages": [
                    {
                        "id": "litchi-eval-id",
                        "name": "litchi-eval",
                        "manifest_path": "/workspace/crates/litchi-eval/Cargo.toml",
                        "features": {},
                        "dependencies": [
                            {"name": "tokio", "kind": "dev"},
                            {"name": "rayon", "kind": "build"},
                            {"name": "reqwest", "kind": None},
                        ],
                    }
                ],
            }
        )

        self.assertEqual(
            snapshot.dependencies["litchi-eval"],
            frozenset({"rayon", "reqwest", "tokio"}),
        )
        self.assertEqual(
            snapshot.normal_dependencies["litchi-eval"],
            frozenset({"reqwest"}),
        )

    def test_xlsb_cannot_depend_on_concrete_xlsx(self) -> None:
        snapshot = valid_snapshot(self.policy)
        edge = boundaries.Edge("litchi-xlsb", "litchi-xlsx")
        edges = dict(snapshot.edges)
        edges[edge] = ("kind=normal, optional=false, target=*, rename=-",)
        dependencies = dict(snapshot.dependencies)
        dependencies[edge.dependent] |= frozenset({edge.dependency})
        snapshot = replace(snapshot, edges=edges, dependencies=dependencies)

        violations = boundaries.audit_snapshot(snapshot, self.policy)

        self.assertIn(
            "OOXML concrete peer edge from litchi-xlsb: litchi-xlsx",
            violations,
        )

    def test_odf_common_cannot_depend_on_format_host(self) -> None:
        snapshot = valid_snapshot(self.policy)
        edge = boundaries.Edge("litchi-odf-common", "litchi-odf")
        edges = dict(snapshot.edges)
        edges[edge] = ("kind=normal, optional=false, target=*, rename=-",)
        dependencies = dict(snapshot.dependencies)
        dependencies[edge.dependent] |= frozenset({edge.dependency})
        snapshot = replace(snapshot, edges=edges, dependencies=dependencies)

        violations = boundaries.audit_snapshot(snapshot, self.policy)

        self.assertIn(
            "foundation crate litchi-odf-common depends upward on: litchi-odf",
            violations,
        )

    def test_odf_common_package_dependencies_are_neutral(self) -> None:
        self.assertEqual(
            self.policy.canonical_edges
            & {
                boundaries.Edge("litchi-odf-common", dependency)
                for dependency in self.policy.packages
            },
            {
                boundaries.Edge("litchi-odf-common", "litchi-core"),
                boundaries.Edge("litchi-odf-common", "soapberry-zip"),
            },
        )

    def test_pptx_drawingml_edge_is_canonical(self) -> None:
        edge = boundaries.Edge("litchi-pptx", "litchi-drawingml")

        self.assertIn(edge, self.policy.canonical_edges)
        self.assertNotIn(edge, self.policy.migration_edges)

    def test_spreadsheet_drawing_owner_and_host_edges_are_canonical(self) -> None:
        owner = "litchi-spreadsheet-drawing"
        dependencies = {
            "litchi-core",
            "litchi-drawingml",
            "litchi-ooxml-common",
            "litchi-opc",
        }

        self.assertEqual(
            {
                edge.dependency
                for edge in self.policy.canonical_edges
                if edge.dependent == owner
            },
            dependencies,
        )
        for host in ("litchi-xlsb", "litchi-xlsx"):
            edge = boundaries.Edge(host, owner)
            self.assertIn(edge, self.policy.canonical_edges)
            self.assertNotIn(edge, self.policy.migration_edges)

    def test_shared_worksheet_view_edges_are_canonical(self) -> None:
        for dependent in ("litchi", "litchi-xlsb", "litchi-xlsx"):
            with self.subTest(dependent=dependent):
                edge = boundaries.Edge(dependent, "litchi-sheet")
                self.assertIn(edge, self.policy.canonical_edges)
                self.assertNotIn(edge, self.policy.migration_edges)

    def test_resolved_migration_edge_requires_policy_cleanup(self) -> None:
        snapshot = valid_snapshot(self.policy)
        edge = boundaries.Edge("litchi-pptx", "litchi-drawingml")
        migration = boundaries.Debt(
            order=1,
            edge=edge,
            reason="test migration",
            exit="test cleanup",
        )
        policy = replace(self.policy, migration_debt=(migration,))
        edges = dict(snapshot.edges)
        del edges[edge]
        dependencies = dict(snapshot.dependencies)
        dependencies[edge.dependent] -= frozenset({edge.dependency})
        snapshot = replace(snapshot, edges=edges, dependencies=dependencies)

        violations = boundaries.audit_snapshot(snapshot, policy)

        self.assertIn(
            f"resolved migration debt still listed: {edge.display()}; remove its policy entry",
            violations,
        )

    def test_resolved_canonical_edge_requires_policy_cleanup(self) -> None:
        snapshot = valid_snapshot(self.policy)
        edge = next(iter(self.policy.canonical_edges))
        edges = dict(snapshot.edges)
        del edges[edge]
        dependencies = dict(snapshot.dependencies)
        dependencies[edge.dependent] -= frozenset({edge.dependency})
        snapshot = replace(snapshot, edges=edges, dependencies=dependencies)

        violations = boundaries.audit_snapshot(snapshot, self.policy)

        self.assertIn(
            f"resolved canonical edge still listed: {edge.display()}; remove its policy entry",
            violations,
        )

    def test_allowed_edges_cannot_hide_a_dependency_cycle(self) -> None:
        snapshot = valid_snapshot(self.policy)
        edge = boundaries.Edge("litchi-core", "litchi-sheet")
        edges = dict(snapshot.edges)
        edges[edge] = ("kind=normal, optional=false, target=*, rename=-",)
        dependencies = dict(snapshot.dependencies)
        dependencies[edge.dependent] |= frozenset({edge.dependency})
        snapshot = replace(snapshot, edges=edges, dependencies=dependencies)
        policy = replace(
            self.policy,
            canonical_edges=self.policy.canonical_edges | frozenset({edge}),
        )

        violations = boundaries.audit_snapshot(snapshot, policy)

        self.assertIn(
            "workspace dependency cycle: litchi-core -> litchi-sheet -> litchi-core",
            violations,
        )

    def test_new_workspace_package_requires_an_inventory_entry(self) -> None:
        snapshot = valid_snapshot(self.policy)
        snapshot = replace(snapshot, packages=snapshot.packages | frozenset({"litchi-new"}))

        violations = boundaries.audit_snapshot(snapshot, self.policy)

        self.assertIn("workspace packages lack topology policy: litchi-new", violations)

    def test_retired_ooxml_monolith_cannot_return(self) -> None:
        raw = copy.deepcopy(self.raw_policy)
        raw["packages"]["litchi-ooxml"] = []

        with self.assertRaisesRegex(
            boundaries.PolicyError,
            "retired monoliths cannot return as workspace packages: litchi-ooxml",
        ):
            boundaries.parse_policy(raw)

    def test_retired_monolith_cannot_return_as_a_policy_package(self) -> None:
        raw = copy.deepcopy(self.raw_policy)
        raw["packages"]["litchi-ole"] = []

        with self.assertRaisesRegex(
            boundaries.PolicyError,
            "retired monoliths cannot return as workspace packages: litchi-ole",
        ):
            boundaries.parse_policy(raw)

    def test_retired_umbrella_feature_cannot_return(self) -> None:
        snapshot = valid_snapshot(self.policy)
        features = dict(snapshot.features)
        features["litchi"] |= frozenset({"full"})
        snapshot = replace(snapshot, features=features)

        violations = boundaries.audit_snapshot(snapshot, self.policy)

        self.assertIn("retired litchi facade features returned: full", violations)

    def test_litchi_facade_rejects_non_optional_normal_dependencies(self) -> None:
        snapshot = valid_snapshot(self.policy)
        normal_dependencies = dict(snapshot.normal_dependencies)
        normal_dependencies["litchi"] |= frozenset({"litchi-xlsb"})
        optional_dependencies = dict(snapshot.normal_optional_dependencies)
        optional_dependencies["litchi"] -= frozenset({"litchi-xlsb"})
        snapshot = replace(
            snapshot,
            normal_dependencies=normal_dependencies,
            normal_optional_dependencies=optional_dependencies,
        )

        violations = boundaries.audit_litchi_facade(snapshot)

        self.assertIn(
            "litchi facade has non-optional normal dependencies: litchi-xlsb",
            violations,
        )

    def test_litchi_facade_rejects_retired_monolithic_iwork_dependency(self) -> None:
        snapshot = valid_snapshot(self.policy)
        dependencies = dict(snapshot.dependencies)
        dependencies["litchi"] |= frozenset({"litchi-iwa"})
        snapshot = replace(snapshot, dependencies=dependencies)

        self.assertIn(
            "litchi facade depends on retired packages: litchi-iwa",
            boundaries.audit_litchi_facade(snapshot),
        )

    def test_litchi_facade_cannot_publicly_reexport_retired_iwa_module(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            path = root / boundaries.FACADE_SOURCE_ROOT / "lib.rs"
            path.parent.mkdir(parents=True)
            path.write_text(
                "\n".join(
                    [
                        "pub mod iwa;",
                        "pub(crate) mod iwa;",
                        "pub use litchi_iwa::Document;",
                        "pub use crate::iwa::Document as LegacyDocument;",
                        "pub(super) use litchi_iwa::{Document, Error};",
                    ]
                )
                + "\n",
                encoding="utf-8",
            )

            self.assertEqual(
                boundaries.audit_litchi_facade_source_topology(root),
                [
                    "retired litchi facade public iwa module: "
                    "crates/litchi/src/lib.rs:1",
                    "retired litchi facade public iwa module: "
                    "crates/litchi/src/lib.rs:2",
                    "retired litchi facade public iwa re-export: "
                    "crates/litchi/src/lib.rs:3",
                    "retired litchi facade public iwa re-export: "
                    "crates/litchi/src/lib.rs:4",
                    "retired litchi facade public iwa re-export: "
                    "crates/litchi/src/lib.rs:5",
                ],
            )

    def test_litchi_facade_iwa_source_policy_ignores_private_and_unrelated_code(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            facade = root / boundaries.FACADE_SOURCE_ROOT / "lib.rs"
            facade.parent.mkdir(parents=True)
            facade.write_text(
                "\n".join(
                    [
                        "// pub mod iwa;",
                        'const NOTE: &str = "pub use litchi_iwa::Document";',
                        "/*",
                        "pub mod iwa;",
                        "/* pub use crate::iwa::Document; */",
                        "pub use litchi_iwa::Document;",
                        "*/",
                        'const RAW_NOTE: &str = r#"',
                        "pub mod iwa;",
                        "pub use litchi_iwa::Document;",
                        '"#;',
                        "mod iwa;",
                        "use litchi_iwa::Document;",
                        "pub mod iwatch;",
                        "pub use crate::iwork::Document;",
                    ]
                )
                + "\n",
                encoding="utf-8",
            )
            focused = root / "crates/litchi-keynote/src/lib.rs"
            focused.parent.mkdir(parents=True)
            focused.write_text(
                "pub mod iwa;\npub use litchi_iwa::Document;\n",
                encoding="utf-8",
            )

            self.assertEqual(boundaries.audit_litchi_facade_source_topology(root), [])

    def test_litchi_facade_requires_an_empty_default_feature(self) -> None:
        snapshot = valid_snapshot(self.policy)
        definitions = dict(snapshot.feature_definitions)
        definitions["litchi"] = dict(definitions["litchi"])
        definitions["litchi"]["default"] = frozenset({"dep:litchi-xlsb"})
        snapshot = replace(snapshot, feature_definitions=definitions)

        self.assertEqual(
            boundaries.audit_litchi_facade(snapshot),
            ["litchi default feature must be exactly empty"],
        )

    def test_litchi_facade_rejects_stale_and_unknown_feature_contracts(self) -> None:
        snapshot = valid_snapshot(self.policy)
        definitions = dict(snapshot.feature_definitions)
        definitions["litchi"] = {
            "default": frozenset(),
            "all": frozenset({"unknown-capability", "dep:litchi-missing"}),
        }
        snapshot = replace(snapshot, feature_definitions=definitions)

        violations = boundaries.audit_litchi_facade(snapshot)

        self.assertEqual(violations, sorted(violations))
        self.assertIn(
            "litchi facade feature references unknown features: "
            "all -> unknown-capability",
            violations,
        )
        self.assertIn(
            "litchi facade has stale dep: feature references: litchi-missing",
            violations,
        )
        self.assertIn("litchi all feature omits optional dependencies: ", violations[0] if violations else "")

    def test_litchi_all_feature_may_cover_optional_dependencies_via_aggregates(self) -> None:
        snapshot = valid_snapshot(self.policy)
        optional_dependencies = snapshot.normal_optional_dependencies["litchi"]
        definitions = dict(snapshot.feature_definitions)
        definitions["litchi"] = {
            "default": frozenset(),
            "all": frozenset({"formats"}),
            "formats": frozenset(
                f"dep:{dependency}" for dependency in optional_dependencies
            ),
        }
        snapshot = replace(snapshot, feature_definitions=definitions)

        self.assertEqual(boundaries.audit_litchi_facade(snapshot), [])

    def test_violations_have_deterministic_order(self) -> None:
        snapshot = valid_snapshot(self.policy)
        snapshot = replace(
            snapshot,
            packages=snapshot.packages | frozenset({"z-new", "a-new"}),
        )

        violations = boundaries.audit_snapshot(snapshot, self.policy)

        self.assertEqual(violations, sorted(violations))
        self.assertIn("workspace packages lack topology policy: a-new, z-new", violations)

    def test_xlsb_host_xlsx_source_cannot_return(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            path = root / "crates/litchi-xlsb/src/host/xlsx/mod.rs"
            path.parent.mkdir(parents=True)
            path.write_text("// retired host\n", encoding="utf-8")

            violations = boundaries.audit_xlsb_source_topology(root)

            self.assertEqual(
                violations,
                [
                    "retired XLSB host XLSX source returned: "
                    "crates/litchi-xlsb/src/host/xlsx"
                ],
            )

    def test_xlsb_package_public_xlsx_module_cannot_return(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            path = root / "crates/litchi-xlsb/src/package/mod.rs"
            path.parent.mkdir(parents=True)
            path.write_text("pub mod xlsx;\n", encoding="utf-8")

            violations = boundaries.audit_xlsb_source_topology(root)

            self.assertEqual(
                violations,
                [
                    "retired XLSB package XLSX module: "
                    "crates/litchi-xlsb/src/package/mod.rs:1"
                ],
            )

    def test_xlsb_package_xlsx_paths_cannot_return(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            path = root / "crates/litchi-xlsb/src/writer.rs"
            path.parent.mkdir(parents=True)
            path.write_text("use crate::package::xlsx::Chart;\n", encoding="utf-8")

            violations = boundaries.audit_xlsb_source_topology(root)

            self.assertEqual(
                violations,
                [
                    "retired XLSB package::xlsx path: "
                    "crates/litchi-xlsb/src/writer.rs:1"
                ],
            )

    def test_retired_shallow_sheet_view_owners_cannot_return(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            for _, retired in boundaries.RETIRED_SHEET_VIEW_OWNER_SOURCES:
                path = root / retired
                path.parent.mkdir(parents=True, exist_ok=True)
                path.write_text("// retired owner\n", encoding="utf-8")
            retired_tree = root / boundaries.RETIRED_XLSX_SHEET_VIEW_OWNER_TREE
            retired_tree.mkdir(parents=True)

            violations = boundaries.audit_spreadsheet_sheet_view_source_topology(root)

            self.assertEqual(
                violations,
                [
                    "retired litchi-xlsb sheet-view owner source returned: "
                    "crates/litchi-xlsb/src/views.rs",
                    "retired litchi-xlsx sheet-view owner source returned: "
                    "crates/litchi-xlsx/src/views.rs",
                    "retired litchi-xlsx sheet-view owner tree returned: "
                    "crates/litchi-xlsx/src/views",
                ],
            )

    def test_retired_iwa_keynote_method_inventory_is_exact(self) -> None:
        self.assertEqual(
            boundaries.RETIRED_IWA_KEYNOTE_METHODS,
            (
                "set_slide_title",
                "replace_slide_title",
                "clear_slide_title",
                "set_slide_body",
                "replace_slide_body",
                "clear_slide_body",
                "set_slide_notes",
                "replace_slide_notes",
                "clear_slide_notes",
                "slide_storage",
                "slide_notes_storage",
            ),
        )

    def test_retired_iwa_keynote_methods_cannot_return(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            path = root / boundaries.IWA_KEYNOTE_SOURCE_ROOT / "legacy/editor.rs"
            path.parent.mkdir(parents=True)
            declarations = [
                ("set_slide_title", "fn r#set_slide_title() {}"),
                ("replace_slide_title", "pub fn replace_slide_title() {}"),
                ("clear_slide_title", "pub(crate) const fn clear_slide_title() {}"),
                ("set_slide_body", "pub(super) async fn set_slide_body() {}"),
                ("replace_slide_body", "pub(in crate) unsafe fn replace_slide_body() {}"),
                ("clear_slide_body", 'pub extern "C" fn clear_slide_body() {}'),
                ("set_slide_notes", "const fn set_slide_notes() {}"),
                ("replace_slide_notes", "async fn replace_slide_notes() {}"),
                ("clear_slide_notes", "unsafe fn clear_slide_notes() {}"),
                ("slide_storage", 'extern "Rust" fn slide_storage() {}'),
                (
                    "slide_notes_storage",
                    "pub async unsafe fn slide_notes_storage() {}",
                ),
            ]
            path.write_text(
                "\n".join(declaration for _, declaration in declarations) + "\n",
                encoding="utf-8",
            )

            violations = boundaries.audit_iwa_keynote_source_topology(root)

            self.assertEqual(
                violations,
                sorted(
                    "retired litchi-iwa Keynote method "
                    f"{name}: crates/litchi-iwa/src/keynote/legacy/editor.rs:{index}"
                    for index, (name, _) in enumerate(declarations, start=1)
                ),
            )

    def test_iwa_keynote_method_policy_matches_only_exact_host_declarations(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            host = root / boundaries.IWA_KEYNOTE_SOURCE_ROOT / "editor.rs"
            host.parent.mkdir(parents=True)
            host.write_text(
                "\n".join(
                    [
                        "// pub fn set_slide_title() {}",
                        'const ADR_NOTE: &str = "fn clear_slide_notes";',
                        "/* pub fn set_slide_body() {}",
                        "   /* fn replace_slide_notes() {} */",
                        "   fn slide_notes_storage() {} */",
                        'const RAW_NOTE: &str = r###"fn clear_slide_title() {}"###;',
                        'const BYTE_RAW_NOTE: &[u8] = br##"fn clear_slide_body() {}"##;',
                        'const C_RAW_NOTE: &CStr = cr#"fn set_slide_notes() {}"#;',
                        'const ESCAPED_NOTE: &str = "\\" fn slide_storage() {}";',
                        "pub fn set_slide_title_text() {}",
                        "pub fn replace_slide_body_text() {}",
                        "pub fn slide_storage_snapshot() {}",
                        "pub fn legacy_set_slide_notes() {}",
                        "pub fn clear_slide_notes_legacy() {}",
                    ]
                )
                + "\n",
                encoding="utf-8",
            )
            non_rust = root / boundaries.IWA_KEYNOTE_SOURCE_ROOT / "retired.txt"
            non_rust.write_text("pub fn set_slide_title() {}\n", encoding="utf-8")
            outside_host_scope = root / "crates/litchi-iwa/src/pages/editor.rs"
            outside_host_scope.parent.mkdir(parents=True)
            outside_host_scope.write_text(
                "pub fn set_slide_title() {}\n",
                encoding="utf-8",
            )
            focused = root / "crates/litchi-keynote/src/package/slide_text.rs"
            focused.parent.mkdir(parents=True)
            focused.write_text(
                "pub fn set_slide_title() {}\npub fn clear_slide_notes() {}\n",
                encoding="utf-8",
            )
            adr = root / "docs/adr/0028-keynote-slide-text.md"
            adr.parent.mkdir(parents=True)
            adr.write_text("Retire `set_slide_title`.\n", encoding="utf-8")

            self.assertEqual(boundaries.audit_iwa_keynote_source_topology(root), [])

    def test_retired_iwa_numbers_table_lock_method_inventory_is_exact(self) -> None:
        self.assertEqual(
            boundaries.RETIRED_IWA_NUMBERS_TABLE_LOCK_METHODS,
            (
                "table_lock_state",
                "set_table_lock_state",
                "table_lock_context",
                "set_table_lock_state_for_model",
                "table_lock_state_for_model",
            ),
        )
        self.assertEqual(
            boundaries.RETIRED_IWA_NUMBERS_TABLE_INFO_FIELDS,
            frozenset({"lock_state"}),
        )

    def test_retired_iwa_numbers_table_lock_methods_cannot_return(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            numbers = root / boundaries.IWA_NUMBERS_SOURCE_ROOT / "legacy/table_lock.rs"
            numbers.parent.mkdir(parents=True)
            numbers_declarations = [
                ("table_lock_state", "fn r#table_lock_state() {}"),
                (
                    "set_table_lock_state",
                    "pub(crate) async unsafe fn set_table_lock_state() {}",
                ),
                (
                    "table_lock_context",
                    'pub(in crate::numbers) const unsafe extern "C" fn '
                    "table_lock_context() {}",
                ),
            ]
            numbers.write_text(
                "\n".join(declaration for _, declaration in numbers_declarations)
                + "\n",
                encoding="utf-8",
            )
            shared = root / boundaries.IWA_TABLE_LOCK_SOURCE
            shared.parent.mkdir(parents=True, exist_ok=True)
            shared.write_text(
                "\n".join(
                    [
                        "pub(crate) const fn r#set_table_lock_state_for_model() {}",
                        'pub unsafe extern "C" fn table_lock_state_for_model() {}',
                    ]
                )
                + "\n",
                encoding="utf-8",
            )

            violations = boundaries.audit_iwa_numbers_table_lock_source_topology(root)

            self.assertEqual(
                violations,
                sorted(
                    [
                        "retired litchi-iwa Numbers table-lock method "
                        f"{name}: crates/litchi-iwa/src/numbers/legacy/table_lock.rs:"
                        f"{index}"
                        for index, (name, _) in enumerate(numbers_declarations, start=1)
                    ]
                    + [
                        "retired litchi-iwa Numbers table-lock method "
                        "set_table_lock_state_for_model: "
                        "crates/litchi-iwa/src/table_lock.rs:1",
                        "retired litchi-iwa Numbers table-lock method "
                        "table_lock_state_for_model: "
                        "crates/litchi-iwa/src/table_lock.rs:2"
                    ]
                ),
            )

    def test_iwa_numbers_table_lock_policy_ignores_non_code_and_near_names(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            numbers = root / boundaries.IWA_NUMBERS_SOURCE_ROOT / "editor/table_lock.rs"
            numbers.parent.mkdir(parents=True)
            numbers.write_text(
                "\n".join(
                    [
                        "// pub fn table_lock_state() {}",
                        'const NOTE: &str = "fn set_table_lock_state() {}";',
                        "/* fn table_lock_context() {}",
                        "   /* pub fn table_lock_state() {} */",
                        "   fn set_table_lock_state() {} */",
                        'const RAW_NOTE: &str = r###"',
                        "fn table_lock_context() {}",
                        '"###;',
                        "pub fn table_lock_state_snapshot() {}",
                        "pub fn reset_table_lock_state() {}",
                        "pub fn table_lock_contextual() {}",
                        "pub fn table_lock_state_for_model() {}",
                        "pub fn set_table_lock_state_for_model() {}",
                    ]
                )
                + "\n",
                encoding="utf-8",
            )
            shared = root / boundaries.IWA_TABLE_LOCK_SOURCE
            shared.write_text(
                "\n".join(
                    [
                        "// fn set_table_lock_state_for_model() {}",
                        'const NOTE: &str = r#"fn set_table_lock_state_for_model() {}"#;',
                        "/* fn table_lock_state_for_model() {} */",
                        'const READ_NOTE: &str = "fn table_lock_state_for_model() {}";',
                        "pub(crate) fn table_lock_state() {}",
                        "pub(crate) fn set_table_lock_state() {}",
                        "fn table_lock_context() {}",
                        "fn set_table_lock_state_for_models() {}",
                    ]
                )
                + "\n",
                encoding="utf-8",
            )

            self.assertEqual(
                boundaries.audit_iwa_numbers_table_lock_source_topology(root), []
            )

    def test_retired_iwa_numbers_table_info_lock_state_field_cannot_return(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            path = root / boundaries.IWA_NUMBERS_TABLE_INFO_SOURCE
            path.parent.mkdir(parents=True)
            path.write_text(
                "\n".join(
                    [
                        "pub struct NumbersTableInfo {",
                        "    pub r#lock_state: LockState,",
                        "}",
                    ]
                )
                + "\n",
                encoding="utf-8",
            )

            self.assertEqual(
                boundaries.audit_iwa_numbers_table_lock_source_topology(root),
                [
                    "retired litchi-iwa Numbers table-info field lock_state: "
                    "crates/litchi-iwa/src/numbers/editor/semantic/model.rs:2"
                ],
            )

    def test_iwa_numbers_table_info_field_policy_ignores_nonpublic_and_other_scopes(
        self,
    ) -> None:
        for permitted_field in (
            "lock_state: LockState,",
            "pub(crate) lock_state: LockState,",
            "pub lock_state_snapshot: LockState,",
        ):
            with self.subTest(permitted_field=permitted_field):
                with tempfile.TemporaryDirectory() as directory:
                    root = Path(directory)
                    path = root / boundaries.IWA_NUMBERS_TABLE_INFO_SOURCE
                    path.parent.mkdir(parents=True)
                    path.write_text(
                        "\n".join(
                            [
                                "/* pub struct NumbersTableInfo {",
                                "    pub lock_state: LockState,",
                                "} */",
                                'const NOTE: &str = r#"',
                                "pub struct NumbersTableInfo {",
                                "    pub lock_state: LockState,",
                                "}",
                                '"#;',
                                "pub struct OtherNumbersTableInfo {",
                                "    pub lock_state: LockState,",
                                "}",
                                "pub struct NumbersTableInfo {",
                                f"    {permitted_field}",
                                "}",
                            ]
                        )
                        + "\n",
                        encoding="utf-8",
                    )
                    focused = root / "crates/litchi-numbers/src/table.rs"
                    focused.parent.mkdir(parents=True)
                    focused.write_text(
                        "pub struct NumbersTableInfo {\n"
                        "    pub lock_state: LockState,\n"
                        "}\n",
                        encoding="utf-8",
                    )
                    pages = root / "crates/litchi-iwa/src/pages/model.rs"
                    pages.parent.mkdir(parents=True)
                    pages.write_text(
                        "pub struct NumbersTableInfo {\n"
                        "    pub lock_state: LockState,\n"
                        "}\n",
                        encoding="utf-8",
                    )

                    self.assertEqual(
                        boundaries.audit_iwa_numbers_table_lock_source_topology(root),
                        [],
                    )

    def test_iwa_numbers_table_lock_policy_ignores_other_owners_and_non_rust_files(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            declarations = "\n".join(
                f"pub fn {name}() {{}}"
                for name in boundaries.RETIRED_IWA_NUMBERS_TABLE_LOCK_METHODS
            ) + "\n"
            for relative in (
                Path("crates/litchi-numbers/src/table/lock.rs"),
                Path("crates/litchi-iwa/src/pages/editor/tables/lock.rs"),
                Path("crates/litchi-iwa/src/keynote/editor/slide_tables/lock.rs"),
            ):
                path = root / relative
                path.parent.mkdir(parents=True, exist_ok=True)
                path.write_text(declarations, encoding="utf-8")
            non_rust = root / boundaries.IWA_NUMBERS_SOURCE_ROOT / "table_lock.txt"
            non_rust.parent.mkdir(parents=True)
            non_rust.write_text(declarations, encoding="utf-8")

            self.assertEqual(
                boundaries.audit_iwa_numbers_table_lock_source_topology(root), []
            )

    def test_retired_iwa_pages_page_layout_method_inventory_is_exact(self) -> None:
        self.assertEqual(
            boundaries.RETIRED_IWA_PAGES_PAGE_LAYOUT_METHODS,
            ("page_layout", "set_page_layout"),
        )

    def test_retired_iwa_pages_page_layout_surface_cannot_return(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            nested = root / boundaries.IWA_PAGES_SOURCE_ROOT / "legacy/layout.rs"
            nested.parent.mkdir(parents=True)
            nested.write_text(
                "\n".join(
                    [
                        "fn r#page_layout() {}",
                        "pub(crate) async unsafe fn set_page_layout() {}",
                    ]
                )
                + "\n",
                encoding="utf-8",
            )
            editor = root / boundaries.IWA_PAGES_EDITOR_SOURCE
            editor.write_text("pub(crate) mod page_layout;\n", encoding="utf-8")
            retired = root / boundaries.RETIRED_IWA_PAGES_PAGE_LAYOUT_SOURCE
            retired.parent.mkdir(parents=True, exist_ok=True)
            retired.write_text("// retired owner returned\n", encoding="utf-8")

            self.assertEqual(
                boundaries.audit_iwa_pages_page_layout_source_topology(root),
                [
                    "retired litchi-iwa Pages page-layout method page_layout: "
                    "crates/litchi-iwa/src/pages/legacy/layout.rs:1",
                    "retired litchi-iwa Pages page-layout method set_page_layout: "
                    "crates/litchi-iwa/src/pages/legacy/layout.rs:2",
                    "retired litchi-iwa Pages page-layout module declaration: "
                    "crates/litchi-iwa/src/pages/editor.rs:1",
                    "retired litchi-iwa Pages page-layout source returned: "
                    "crates/litchi-iwa/src/pages/editor/page_layout.rs",
                ],
            )

    def test_retired_iwa_pages_page_layout_module_declaration_variants(
        self,
    ) -> None:
        for declaration in (
            "mod page_layout;",
            "pub(crate) mod page_layout;",
            "pub mod r#page_layout;",
            "mod\npage_layout\n{}",
        ):
            with self.subTest(declaration=declaration):
                with tempfile.TemporaryDirectory() as directory:
                    root = Path(directory)
                    editor = root / boundaries.IWA_PAGES_EDITOR_SOURCE
                    editor.parent.mkdir(parents=True)
                    editor.write_text(declaration + "\n", encoding="utf-8")

                    self.assertEqual(
                        boundaries.audit_iwa_pages_page_layout_source_topology(root),
                        [
                            "retired litchi-iwa Pages page-layout module declaration: "
                            "crates/litchi-iwa/src/pages/editor.rs:1"
                        ],
                    )

    def test_iwa_pages_page_layout_policy_ignores_non_code_near_names_and_owners(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            host = root / boundaries.IWA_PAGES_SOURCE_ROOT / "legacy/page_layout_old.rs"
            host.parent.mkdir(parents=True)
            host.write_text(
                "\n".join(
                    [
                        "// pub fn page_layout() {}",
                        'const NOTE: &str = "fn set_page_layout() {}";',
                        "/* fn page_layout() {}",
                        "   /* fn set_page_layout() {} */",
                        "   fn page_layout() {} */",
                        'const RAW_NOTE: &str = r###"fn set_page_layout() {}"###;',
                        "pub fn page_layout_snapshot() {}",
                        "pub fn reset_page_layout() {}",
                        "pub fn set_page_layouts() {}",
                    ]
                )
                + "\n",
                encoding="utf-8",
            )
            editor = root / boundaries.IWA_PAGES_EDITOR_SOURCE
            editor.write_text(
                "\n".join(
                    [
                        "// mod page_layout;",
                        'const NOTE: &str = "mod page_layout;";',
                        "/* mod page_layout; */",
                        "mod page_layout_legacy;",
                        "use crate::page_layout;",
                    ]
                )
                + "\n",
                encoding="utf-8",
            )
            declarations = "pub fn page_layout() {}\npub fn set_page_layout() {}\n"
            for relative in (
                Path("crates/litchi-pages/src/package/page_layout.rs"),
                Path("crates/litchi-iwa/src/keynote/editor/page_layout.rs"),
            ):
                path = root / relative
                path.parent.mkdir(parents=True, exist_ok=True)
                path.write_text(declarations, encoding="utf-8")
            non_rust = root / boundaries.IWA_PAGES_SOURCE_ROOT / "page_layout.txt"
            non_rust.write_text(declarations, encoding="utf-8")

            self.assertEqual(
                boundaries.audit_iwa_pages_page_layout_source_topology(root), []
            )

    def test_focused_pages_page_layout_public_api_rejects_physical_leaks(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            implementation = root / boundaries.PAGES_PAGE_LAYOUT_IMPLEMENTATION_SOURCE
            implementation.parent.mkdir(parents=True)
            implementation.write_text(
                "\n".join(
                    [
                        "pub async unsafe fn r#page_layout(",
                        "    r#object_id: u64,",
                        ") {}",
                        "pub type PageLayoutPatch = DocumentArchive;",
                        "pub type PageLayoutCommit = IWorkPackage;",
                        "pub type PageLayoutDiagnostics = buffa::DocumentArchiveView;",
                        "impl prost::Message for PageLayoutEdit {}",
                    ]
                )
                + "\n",
                encoding="utf-8",
            )
            semantic = root / boundaries.PAGES_PAGE_LAYOUT_SEMANTIC_SOURCE
            semantic.write_text(
                "pub fn page_layout(source_bytes: &[u8]) -> SourceBytes {}\n",
                encoding="utf-8",
            )
            lib_export, package_export = (
                root / path for path in boundaries.PAGES_PAGE_LAYOUT_EXPORT_SOURCES
            )
            package_export.parent.mkdir(parents=True, exist_ok=True)
            package_export.write_text(
                "pub type PageLayoutDiagnostics = litchi_iwa_core::RawObject;\n",
                encoding="utf-8",
            )
            lib_export.write_text(
                "pub use litchi_iwa_protos::PageLayoutArchive as PageLayoutPatch;\n",
                encoding="utf-8",
            )

            violations = boundaries.audit_pages_page_layout_facade_source_topology(root)

            self.assertEqual(violations, sorted(violations))
            self.assertEqual(
                violations,
                sorted(
                    [
                        "focused litchi-pages page-layout public API exposes "
                        "raw identifier object_id: "
                        "crates/litchi-pages/src/package/page_layout.rs:2",
                        "focused litchi-pages page-layout public API exposes "
                        "archive/IWA type DocumentArchive: "
                        "crates/litchi-pages/src/package/page_layout.rs:4",
                        "focused litchi-pages page-layout public API exposes "
                        "archive/IWA type IWorkPackage: "
                        "crates/litchi-pages/src/package/page_layout.rs:5",
                        "focused litchi-pages page-layout public API exposes "
                        "protobuf type buffa: "
                        "crates/litchi-pages/src/package/page_layout.rs:6",
                        "focused litchi-pages page-layout public API exposes "
                        "archive/IWA type DocumentArchiveView: "
                        "crates/litchi-pages/src/package/page_layout.rs:6",
                        "focused litchi-pages page-layout public API exposes "
                        "protobuf type prost: "
                        "crates/litchi-pages/src/package/page_layout.rs:7",
                        "focused litchi-pages page-layout public API exposes "
                        "protobuf type Message: "
                        "crates/litchi-pages/src/package/page_layout.rs:7",
                        "focused litchi-pages page-layout public API exposes "
                        "archive/IWA type litchi_iwa_core: "
                        "crates/litchi-pages/src/package.rs:1",
                        "focused litchi-pages page-layout public API exposes "
                        "native object RawObject: "
                        "crates/litchi-pages/src/package.rs:1",
                        "focused litchi-pages page-layout public API exposes "
                        "archive/IWA type litchi_iwa_protos: "
                        "crates/litchi-pages/src/lib.rs:1",
                        "focused litchi-pages page-layout public API exposes "
                        "archive/IWA type PageLayoutArchive: "
                        "crates/litchi-pages/src/lib.rs:1",
                        "focused litchi-pages page-layout public API exposes "
                        "raw source bytes source_bytes: "
                        "crates/litchi-pages/src/page_layout.rs:1",
                        "focused litchi-pages page-layout public API exposes "
                        "raw byte slice &[u8]: "
                        "crates/litchi-pages/src/page_layout.rs:1",
                        "focused litchi-pages page-layout public API exposes "
                        "raw source bytes SourceBytes: "
                        "crates/litchi-pages/src/page_layout.rs:1",
                    ]
                ),
            )

    def test_focused_pages_page_layout_public_api_ignores_safe_and_private_code(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            implementation = root / boundaries.PAGES_PAGE_LAYOUT_IMPLEMENTATION_SOURCE
            implementation.parent.mkdir(parents=True)
            implementation.write_text(
                "\n".join(
                    [
                        "// pub fn page_layout(object_id: ObjectId) {}",
                        'const NOTE: &str = "pub type PageLayoutPatch = DocumentArchive";',
                        "/* pub type PageLayoutCommit = IWorkPackage;",
                        "   /* impl buffa::Message for PageLayoutEdit {} */",
                        "*/",
                        'const RAW_NOTE: &str = r###"',
                        "pub fn page_layout(object_id: ObjectId) {}",
                        '"###;',
                        "pub struct PageLayoutEdit;",
                        "pub fn page_layout(layout: Layout) "
                        "-> Result<PageLayoutCommit, PageLayoutError> {}",
                        "impl PageLayoutEdit {}",
                        "impl prost::Message for Unrelated {}",
                        "fn private_page_layout(object_id: u64) {}",
                        "pub(crate) fn restricted_page_layout(archive: DocumentArchive) {}",
                    ]
                )
                + "\n",
                encoding="utf-8",
            )
            lib_export, package_export = (
                root / path for path in boundaries.PAGES_PAGE_LAYOUT_EXPORT_SOURCES
            )
            package_export.parent.mkdir(parents=True, exist_ok=True)
            safe_exports = (
                "pub use crate::page_layout::Layout;\n"
                "pub use package::page_layout::{PageLayoutCommit, PageLayoutError};\n"
                "pub fn unrelated(object_id: u64) -> DocumentArchive { todo!() }\n"
                "pub type PageLayoutsCommit = litchi_iwa_core::RawObject;\n"
            )
            package_export.write_text(safe_exports, encoding="utf-8")
            lib_export.write_text(safe_exports, encoding="utf-8")
            low_level = root / boundaries.PAGES_SOURCE_ROOT / "page_layout.rs"
            low_level.write_text(
                "\n".join(
                    [
                        "pub fn page_layout(input: InputBytes, byte_count: usize) "
                        "-> OutputBytes { todo!() }",
                        "fn private_source_bytes(source_bytes: &[u8]) -> SourceBytes "
                        "{ todo!() }",
                        "pub(crate) fn restricted_source_bytes(source_bytes: &[u8]) "
                        "-> SourceBytes { todo!() }",
                    ]
                )
                + "\n",
                encoding="utf-8",
            )
            unrelated_owner = root / "crates/litchi-keynote/src/package/page_layout.rs"
            unrelated_owner.parent.mkdir(parents=True)
            unrelated_owner.write_text(
                "pub type PageLayoutPatch = litchi_iwa_core::RawObject;\n",
                encoding="utf-8",
            )

            self.assertEqual(
                boundaries.audit_pages_page_layout_facade_source_topology(root), []
            )

    def test_retired_iwa_pages_document_settings_method_inventory_is_exact(
        self,
    ) -> None:
        self.assertEqual(
            boundaries.RETIRED_IWA_PAGES_DOCUMENT_SETTINGS_METHODS,
            (
                "document_options",
                "set_document_options",
                "footnote_settings",
                "set_footnote_settings",
            ),
        )
        self.assertEqual(
            boundaries.RETIRED_IWA_PAGES_DOCUMENT_SETTINGS_MODULES,
            ("document_options", "footnote_settings"),
        )

    def test_retired_iwa_pages_document_settings_surface_cannot_return(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            nested = root / boundaries.IWA_PAGES_SOURCE_ROOT / "legacy/settings.rs"
            nested.parent.mkdir(parents=True)
            declarations = [
                ("document_options", "fn r#document_options() {}"),
                (
                    "set_document_options",
                    "pub(crate) async unsafe fn set_document_options() {}",
                ),
                (
                    "footnote_settings",
                    'pub(in crate::pages) const unsafe extern "C" fn '
                    "footnote_settings() {}",
                ),
                (
                    "set_footnote_settings",
                    "pub(super) fn r#set_footnote_settings() {}",
                ),
            ]
            nested.write_text(
                "\n".join(declaration for _, declaration in declarations) + "\n",
                encoding="utf-8",
            )
            editor = root / boundaries.IWA_PAGES_EDITOR_SOURCE
            editor.write_text(
                "pub(crate) mod r#document_options;\n"
                "pub mod footnote_settings {}\n",
                encoding="utf-8",
            )
            for retired in boundaries.RETIRED_IWA_PAGES_DOCUMENT_SETTINGS_SOURCES:
                path = root / retired
                path.parent.mkdir(parents=True, exist_ok=True)
                path.write_text("// retired owner returned\n", encoding="utf-8")

            self.assertEqual(
                boundaries.audit_iwa_pages_document_settings_source_topology(root),
                sorted(
                    [
                        "retired litchi-iwa Pages document-settings method "
                        f"{name}: crates/litchi-iwa/src/pages/legacy/settings.rs:"
                        f"{index}"
                        for index, (name, _) in enumerate(declarations, start=1)
                    ]
                    + [
                        "retired litchi-iwa Pages document-settings module "
                        "document_options: crates/litchi-iwa/src/pages/editor.rs:1",
                        "retired litchi-iwa Pages document-settings module "
                        "footnote_settings: crates/litchi-iwa/src/pages/editor.rs:2",
                    ]
                    + [
                        "retired litchi-iwa Pages document-settings source returned: "
                        f"{retired}"
                        for retired in boundaries.RETIRED_IWA_PAGES_DOCUMENT_SETTINGS_SOURCES
                    ]
                ),
            )

    def test_retired_iwa_pages_document_settings_module_declaration_variants(
        self,
    ) -> None:
        for module in boundaries.RETIRED_IWA_PAGES_DOCUMENT_SETTINGS_MODULES:
            for declaration in (
                f"mod {module};",
                f"pub(crate) mod {module};",
                f"pub mod r#{module};",
                f"mod\n{module}\n{{}}",
            ):
                with self.subTest(module=module, declaration=declaration):
                    with tempfile.TemporaryDirectory() as directory:
                        root = Path(directory)
                        editor = root / boundaries.IWA_PAGES_EDITOR_SOURCE
                        editor.parent.mkdir(parents=True)
                        editor.write_text(declaration + "\n", encoding="utf-8")

                        self.assertEqual(
                            boundaries.audit_iwa_pages_document_settings_source_topology(
                                root
                            ),
                            [
                                "retired litchi-iwa Pages document-settings module "
                                f"{module}: crates/litchi-iwa/src/pages/editor.rs:1"
                            ],
                        )

    def test_iwa_pages_document_settings_policy_ignores_trivia_near_names_and_owners(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            host = root / boundaries.IWA_PAGES_SOURCE_ROOT / "legacy/settings_old.rs"
            host.parent.mkdir(parents=True)
            host.write_text(
                "\n".join(
                    [
                        "// pub fn document_options() {}",
                        'const NOTE: &str = "fn set_document_options() {}";',
                        "/* fn footnote_settings() {}",
                        "   /* fn set_footnote_settings() {} */",
                        "   fn document_options() {} */",
                        'const RAW_NOTE: &str = r###"fn footnote_settings() {}"###;',
                        "pub fn document_options_snapshot() {}",
                        "pub fn reset_document_options() {}",
                        "pub fn footnote_settings_snapshot() {}",
                        "pub fn set_footnote_settings_for_section() {}",
                    ]
                )
                + "\n",
                encoding="utf-8",
            )
            editor = root / boundaries.IWA_PAGES_EDITOR_SOURCE
            editor.write_text(
                "\n".join(
                    [
                        "// mod document_options;",
                        'const NOTE: &str = "mod footnote_settings;";',
                        "/* mod document_options; */",
                        'const RAW_NOTE: &str = r#"mod footnote_settings;"#;',
                        "mod document_options_legacy;",
                        "mod footnote_settings_legacy;",
                        "use crate::document_options;",
                    ]
                )
                + "\n",
                encoding="utf-8",
            )
            declarations = "\n".join(
                f"pub fn {name}() {{}}"
                for name in boundaries.RETIRED_IWA_PAGES_DOCUMENT_SETTINGS_METHODS
            ) + "\n"
            for relative in (
                Path("crates/litchi-pages/src/package/document_settings.rs"),
                Path("crates/litchi-iwa/src/keynote/editor/document_settings.rs"),
            ):
                path = root / relative
                path.parent.mkdir(parents=True, exist_ok=True)
                path.write_text(declarations, encoding="utf-8")
            non_rust = root / boundaries.IWA_PAGES_SOURCE_ROOT / "document_options.txt"
            non_rust.write_text(declarations, encoding="utf-8")

            self.assertEqual(
                boundaries.audit_iwa_pages_document_settings_source_topology(root), []
            )

    def test_focused_pages_document_settings_public_api_rejects_physical_leaks(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            semantic, transaction = (
                root / path
                for path in boundaries.PAGES_DOCUMENT_SETTINGS_IMPLEMENTATION_SOURCES
            )
            semantic.parent.mkdir(parents=True)
            semantic.write_text(
                "\n".join(
                    [
                        "pub fn document_settings(source_bytes: &[u8], object_id: u64) "
                        "-> DocumentArchive {}",
                        "pub type Patch = buffa::DocumentArchiveView;",
                        "impl prost::Message for document_settings::Edit {}",
                    ]
                )
                + "\n",
                encoding="utf-8",
            )
            transaction.parent.mkdir(parents=True, exist_ok=True)
            transaction.write_text(
                "\n".join(
                    [
                        "pub type Commit = IWorkPackage;",
                        "pub type Diagnostics = litchi_iwa_core::RawObject;",
                        "pub type Error = SourceBytes;",
                    ]
                )
                + "\n",
                encoding="utf-8",
            )
            lib_export, package_export = (
                root / path for path in boundaries.PAGES_DOCUMENT_SETTINGS_EXPORT_SOURCES
            )
            lib_export.write_text(
                "pub use litchi_iwa_protos::document_settings::GeneratedArchive;\n",
                encoding="utf-8",
            )
            package_export.write_text(
                "pub type document_settings_limit = prost_types::MessageInfo;\n",
                encoding="utf-8",
            )

            violations = boundaries.audit_pages_document_settings_facade_source_topology(
                root
            )

            self.assertEqual(violations, sorted(violations))
            self.assertTrue(
                all(
                    violation.startswith(
                        "focused litchi-pages document-settings public API exposes "
                    )
                    for violation in violations
                )
            )
            expected_fragments = (
                "raw source bytes source_bytes",
                "raw byte slice &[u8]",
                "raw identifier object_id",
                "archive/IWA type DocumentArchive",
                "protobuf type buffa",
                "archive/IWA type DocumentArchiveView",
                "protobuf type prost",
                "protobuf type Message",
                "archive/IWA type IWorkPackage",
                "archive/IWA type litchi_iwa_core",
                "native object RawObject",
                "raw source bytes SourceBytes",
                "archive/IWA type litchi_iwa_protos",
                "archive/IWA type GeneratedArchive",
                "protobuf type prost_types",
                "archive/IWA type MessageInfo",
            )
            self.assertEqual(len(violations), len(expected_fragments))
            for fragment in expected_fragments:
                with self.subTest(fragment=fragment):
                    self.assertTrue(
                        any(fragment in violation for violation in violations),
                        msg=f"missing focused document-settings leak: {fragment}",
                    )

    def test_focused_pages_document_settings_api_ignores_semantic_and_private_code(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            semantic, transaction = (
                root / path
                for path in boundaries.PAGES_DOCUMENT_SETTINGS_IMPLEMENTATION_SOURCES
            )
            semantic.parent.mkdir(parents=True)
            semantic.write_text(
                "\n".join(
                    [
                        "// pub fn document_settings(object_id: ObjectId) {}",
                        'const NOTE: &str = "pub type DocumentSettingsPatch = DocumentArchive";',
                        "/* impl buffa::Message for DocumentSettingsEdit {} */",
                        "// pub struct DocumentSettings;",
                        'const FLAT_NOTE: &str = "pub type DocumentSettings = Settings";',
                        "pub struct Options;",
                        "pub struct FootnoteSettings;",
                        "pub struct Settings;",
                        "pub struct DocumentSettingsSnapshot;",
                        "pub fn document_settings(options: Options, footnotes: FootnoteSettings, "
                        "input: InputBytes, byte_count: usize) -> OutputBytes { todo!() }",
                        "fn private_source_bytes(source_bytes: &[u8], object_id: u64) "
                        "-> SourceBytes { todo!() }",
                        "pub(crate) fn restricted_settings(archive: DocumentArchive) {}",
                        "struct DocumentSettings;",
                        "pub(crate) struct DocumentSettings;",
                        "struct DocumentSettingsEdit;",
                        "pub(crate) struct DocumentSettingsPatch;",
                        "impl prost::Message for Unrelated {}",
                    ]
                )
                + "\n",
                encoding="utf-8",
            )
            transaction.parent.mkdir(parents=True, exist_ok=True)
            transaction.write_text(
                "\n".join(
                    [
                        "pub struct Edit;",
                        "pub struct Patch;",
                        "pub struct Commit;",
                        "pub struct Diagnostics;",
                        "pub struct Error;",
                        "pub struct LimitKind;",
                        "impl Edit {}",
                    ]
                )
                + "\n",
                encoding="utf-8",
            )
            lib_export, package_export = (
                root / path for path in boundaries.PAGES_DOCUMENT_SETTINGS_EXPORT_SOURCES
            )
            safe_exports = (
                "pub use crate::document_settings::{Options, FootnoteSettings, Settings};\n"
                "pub use package::document_settings::{Edit, Patch, Commit, Diagnostics, "
                "Error, LimitKind};\n"
                "pub fn unrelated(object_id: u64) -> DocumentArchive { todo!() }\n"
                "pub type DocumentSettingCommit = litchi_iwa_core::RawObject;\n"
            )
            lib_export.write_text(safe_exports, encoding="utf-8")
            package_export.write_text(safe_exports, encoding="utf-8")
            other_owner = root / "crates/litchi-keynote/src/document_settings.rs"
            other_owner.parent.mkdir(parents=True)
            other_owner.write_text(
                "pub struct DocumentSettings;\n"
                "pub type DocumentSettingsPatch = litchi_iwa_core::RawObject;\n",
                encoding="utf-8",
            )

            self.assertEqual(
                boundaries.audit_pages_document_settings_facade_source_topology(root), []
            )

    def test_focused_pages_document_settings_public_api_rejects_flat_aliases(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            semantic, transaction = (
                root / path
                for path in boundaries.PAGES_DOCUMENT_SETTINGS_IMPLEMENTATION_SOURCES
            )
            semantic.parent.mkdir(parents=True)
            semantic.write_text(
                "pub struct DocumentSettings;\n"
                "pub struct DocumentSettingsEdit;\n"
                "pub type DocumentSettingsPatch = Patch;\n",
                encoding="utf-8",
            )
            transaction.parent.mkdir(parents=True, exist_ok=True)
            transaction.write_text(
                "pub type DocumentSettingsCommit = Commit;\n"
                "pub type DocumentSettingsDiagnostics = Diagnostics;\n",
                encoding="utf-8",
            )
            lib_export, package_export = (
                root / path for path in boundaries.PAGES_DOCUMENT_SETTINGS_EXPORT_SOURCES
            )
            lib_export.write_text(
                "pub use crate::document_settings::Error as DocumentSettingsError;\n",
                encoding="utf-8",
            )
            package_export.write_text(
                "pub use package::document_settings::LimitKind "
                "as DocumentSettingsLimitKind;\n",
                encoding="utf-8",
            )

            self.assertEqual(
                boundaries.audit_pages_document_settings_facade_source_topology(root),
                sorted(
                    [
                        "focused litchi-pages document-settings public API retains "
                        "flat alias DocumentSettings: "
                        "crates/litchi-pages/src/document_settings.rs:1",
                        "focused litchi-pages document-settings public API retains "
                        "flat alias DocumentSettingsEdit: "
                        "crates/litchi-pages/src/document_settings.rs:2",
                        "focused litchi-pages document-settings public API retains "
                        "flat alias DocumentSettingsPatch: "
                        "crates/litchi-pages/src/document_settings.rs:3",
                        "focused litchi-pages document-settings public API retains "
                        "flat alias DocumentSettingsCommit: "
                        "crates/litchi-pages/src/package/document_settings.rs:1",
                        "focused litchi-pages document-settings public API retains "
                        "flat alias DocumentSettingsDiagnostics: "
                        "crates/litchi-pages/src/package/document_settings.rs:2",
                        "focused litchi-pages document-settings public API retains "
                        "flat alias DocumentSettingsError: "
                        "crates/litchi-pages/src/lib.rs:1",
                        "focused litchi-pages document-settings public API retains "
                        "flat alias DocumentSettingsLimitKind: "
                        "crates/litchi-pages/src/package.rs:1",
                    ]
                ),
            )

    def test_legacy_xlsb_sheet_view_names_and_methods_are_forbidden(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            path = root / boundaries.XLSB_SOURCE_ROOT / "legacy_sheet_view.rs"
            path.parent.mkdir(parents=True)
            path.write_text(
                "\n".join(
                    [
                        *boundaries.LEGACY_XLSB_SHEET_VIEW_NAMES,
                        "pub fn set_sheet_view() {}",
                        "pub(crate) fn sheet_views() {}",
                    ]
                )
                + "\n",
                encoding="utf-8",
            )

            violations = boundaries.audit_spreadsheet_sheet_view_source_topology(root)

            self.assertEqual(violations, sorted(violations))
            self.assertEqual(
                violations,
                [
                    f"litchi-xlsb legacy sheet-view name {name}: "
                    "crates/litchi-xlsb/src/legacy_sheet_view.rs:"
                    f"{index}"
                    for index, name in enumerate(
                        boundaries.LEGACY_XLSB_SHEET_VIEW_NAMES, start=1
                    )
                ]
                + [
                    "litchi-xlsb legacy sheet-view public method set_sheet_view: "
                    "crates/litchi-xlsb/src/legacy_sheet_view.rs:7",
                    "litchi-xlsb legacy sheet-view public method sheet_views: "
                    "crates/litchi-xlsb/src/legacy_sheet_view.rs:8",
                ],
            )

    def test_sheet_view_hosts_cannot_define_canonical_view_types(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            for path in (
                boundaries.XLSB_SHEET_VIEW_ADAPTER,
                boundaries.XLSX_SHEET_VIEW_MODEL,
            ):
                absolute_path = root / path
                absolute_path.parent.mkdir(parents=True, exist_ok=True)
                absolute_path.write_text(
                    "\n".join(
                        [
                            *(f"pub struct {name};" for name in boundaries.CANONICAL_SHEET_VIEW_TYPES),
                            "pub struct Entry;",
                            "pub struct Collection;",
                            "pub struct PivotSelection;",
                            "pub struct Extension;",
                        ]
                    )
                    + "\n",
                    encoding="utf-8",
                )

            violations = boundaries.audit_spreadsheet_sheet_view_source_topology(root)

            self.assertEqual(violations, sorted(violations))
            self.assertEqual(
                violations,
                [
                    f"{host} sheet-view {role} defines canonical view type {name}: "
                    f"{path}:{index}"
                    for host, path, role in (
                        ("litchi-xlsb", boundaries.XLSB_SHEET_VIEW_ADAPTER, "adapter"),
                        ("litchi-xlsx", boundaries.XLSX_SHEET_VIEW_MODEL, "model"),
                    )
                    for index, name in enumerate(
                        boundaries.CANONICAL_SHEET_VIEW_TYPES, start=1
                    )
                ],
            )

    def test_retired_xlsx_chart_owner_sources_cannot_return(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            for retired in boundaries.RETIRED_XLSX_CHART_FILES:
                path = root / boundaries.XLSX_SOURCE_ROOT / retired
                path.parent.mkdir(parents=True, exist_ok=True)
                path.write_text("// retired owner\n", encoding="utf-8")

            violations = boundaries.audit_spreadsheet_chart_source_topology(root)

            self.assertEqual(
                violations,
                [
                    "retired XLSX chart owner source returned: "
                    f"{boundaries.XLSX_SOURCE_ROOT / retired}"
                    for retired in boundaries.RETIRED_XLSX_CHART_FILES
                ],
            )

    def test_spreadsheet_chart_facades_must_remain_thin(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            for facades in boundaries.SPREADSHEET_CHART_FACADES.values():
                path = root / facades[0]
                path.parent.mkdir(parents=True, exist_ok=True)
                path.write_text(
                    "// facade\n"
                    * (boundaries.MAX_SPREADSHEET_CHART_FACADE_LINES + 1),
                    encoding="utf-8",
                )

            violations = boundaries.audit_spreadsheet_chart_source_topology(root)

            self.assertEqual(
                violations,
                [
                    "litchi-xlsb chart facade exceeds "
                    f"{boundaries.MAX_SPREADSHEET_CHART_FACADE_LINES} lines: "
                    "crates/litchi-xlsb/src/chart.rs",
                    "litchi-xlsx chart facade exceeds "
                    f"{boundaries.MAX_SPREADSHEET_CHART_FACADE_LINES} lines: "
                    "crates/litchi-xlsx/src/chart.rs",
                ],
            )

    def test_chart_facades_cannot_define_shared_chart_types(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            for host, facades in boundaries.SPREADSHEET_CHART_FACADES.items():
                path = root / facades[0]
                path.parent.mkdir(parents=True, exist_ok=True)
                path.write_text("pub struct Chart;\n", encoding="utf-8")

            violations = boundaries.audit_spreadsheet_chart_source_topology(root)

            self.assertEqual(
                violations,
                [
                    "litchi-xlsb chart facade defines shared chart type: "
                    "crates/litchi-xlsb/src/chart.rs:1",
                    "litchi-xlsx chart facade defines shared chart type: "
                    "crates/litchi-xlsx/src/chart.rs:1",
                ],
            )

    def test_chart_facades_cannot_directly_use_drawingml_chart_codecs(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            for host, facades in boundaries.SPREADSHEET_CHART_FACADES.items():
                path = root / facades[0]
                path.parent.mkdir(parents=True, exist_ok=True)
                path.write_text(
                    "litchi_drawingml::chart::read_chart();\n", encoding="utf-8"
                )

            violations = boundaries.audit_spreadsheet_chart_source_topology(root)

            self.assertEqual(
                violations,
                [
                    "litchi-xlsb chart facade directly uses litchi_drawingml chart codec: "
                    "crates/litchi-xlsb/src/chart.rs:1",
                    "litchi-xlsx chart facade directly uses litchi_drawingml chart codec: "
                    "crates/litchi-xlsx/src/chart.rs:1",
                ],
            )

    def test_retired_xlsx_shape_owner_sources_cannot_return(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            for retired in boundaries.RETIRED_XLSX_SHAPE_FILES:
                path = root / boundaries.XLSX_SOURCE_ROOT / retired
                path.parent.mkdir(parents=True, exist_ok=True)
                path.write_text("// retired owner\n", encoding="utf-8")

            violations = boundaries.audit_spreadsheet_shape_source_topology(root)

            self.assertEqual(
                violations,
                [
                    "retired XLSX shape owner source returned: "
                    f"{boundaries.XLSX_SOURCE_ROOT / retired}"
                    for retired in boundaries.RETIRED_XLSX_SHAPE_FILES
                ],
            )

    def test_spreadsheet_shape_facades_must_remain_thin(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            for facades in boundaries.SPREADSHEET_SHAPE_FACADES.values():
                path = root / facades[0]
                path.parent.mkdir(parents=True, exist_ok=True)
                path.write_text(
                    "// facade\n"
                    * (boundaries.MAX_SPREADSHEET_SHAPE_FACADE_LINES + 1),
                    encoding="utf-8",
                )

            violations = boundaries.audit_spreadsheet_shape_source_topology(root)

            self.assertEqual(
                violations,
                [
                    "litchi-xlsb shape facade exceeds "
                    f"{boundaries.MAX_SPREADSHEET_SHAPE_FACADE_LINES} lines: "
                    "crates/litchi-xlsb/src/shapes.rs",
                    "litchi-xlsx shape facade exceeds "
                    f"{boundaries.MAX_SPREADSHEET_SHAPE_FACADE_LINES} lines: "
                    "crates/litchi-xlsx/src/shapes/mod.rs",
                ],
            )

    def test_shape_facades_cannot_define_local_types(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            for host, facades in boundaries.SPREADSHEET_SHAPE_FACADES.items():
                path = root / facades[0]
                path.parent.mkdir(parents=True, exist_ok=True)
                path.write_text("pub struct Shape;\n", encoding="utf-8")

            violations = boundaries.audit_spreadsheet_shape_source_topology(root)

            self.assertEqual(
                violations,
                [
                    "litchi-xlsb shape facade defines local shape type: "
                    "crates/litchi-xlsb/src/shapes.rs:1",
                    "litchi-xlsx shape facade defines local shape type: "
                    "crates/litchi-xlsx/src/shapes/mod.rs:1",
                ],
            )

    def test_shape_facades_cannot_directly_use_xml_implementation_details(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            for facades in boundaries.SPREADSHEET_SHAPE_FACADES.values():
                path = root / facades[0]
                path.parent.mkdir(parents=True, exist_ok=True)
                path.write_text(
                    "use quick_xml::Reader;\nwrite(\"<xdr:wsDr>\");\n",
                    encoding="utf-8",
                )

            violations = boundaries.audit_spreadsheet_shape_source_topology(root)

            self.assertEqual(
                violations,
                [
                    "litchi-xlsb shape facade directly emits xdr XML: "
                    "crates/litchi-xlsb/src/shapes.rs:2",
                    "litchi-xlsb shape facade directly uses quick_xml: "
                    "crates/litchi-xlsb/src/shapes.rs:1",
                    "litchi-xlsx shape facade directly emits xdr XML: "
                    "crates/litchi-xlsx/src/shapes/mod.rs:2",
                    "litchi-xlsx shape facade directly uses quick_xml: "
                    "crates/litchi-xlsx/src/shapes/mod.rs:1",
                ],
            )

    def test_legacy_host_shape_names_are_forbidden(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            for host, source_root in (
                ("litchi-xlsb", boundaries.XLSB_SOURCE_ROOT),
                ("litchi-xlsx", boundaries.XLSX_SOURCE_ROOT),
            ):
                path = root / source_root / "shape_host.rs"
                path.parent.mkdir(parents=True, exist_ok=True)
                path.write_text(
                    "\n".join(boundaries.LEGACY_HOST_SHAPE_NAMES) + "\n",
                    encoding="utf-8",
                )

            violations = boundaries.audit_spreadsheet_shape_source_topology(root)

            self.assertEqual(
                violations,
                [
                    f"{host} legacy shape host name {name}: "
                    f"{source_root}/shape_host.rs:{index}"
                    for host, source_root in (
                        ("litchi-xlsb", boundaries.XLSB_SOURCE_ROOT),
                        ("litchi-xlsx", boundaries.XLSX_SOURCE_ROOT),
                    )
                    for index, name in enumerate(boundaries.LEGACY_HOST_SHAPE_NAMES, start=1)
                ],
            )


if __name__ == "__main__":
    unittest.main()
