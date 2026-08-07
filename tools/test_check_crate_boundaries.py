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
    for edge in edges:
        dependencies[edge.dependent].add(edge.dependency)
    dependencies["litchi-core"].update(item.name for item in policy.core_dependency_debt)
    frozen_dependencies = {
        name: frozenset(items) for name, items in dependencies.items()
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
        edges={edge: ("kind=normal, optional=false, target=*, rename=-",) for edge in edges},
        dependencies=frozen_dependencies,
        normal_dependencies=frozen_dependencies,
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
                "pub mod iwa;\npub use litchi_iwa::Document;\n",
                encoding="utf-8",
            )

            self.assertEqual(
                boundaries.audit_litchi_facade_source_topology(root),
                [
                    "retired litchi facade public iwa module: "
                    "crates/litchi/src/lib.rs:1",
                    "retired litchi facade public iwa re-export: "
                    "crates/litchi/src/lib.rs:2",
                ],
            )

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
