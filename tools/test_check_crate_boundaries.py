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


def add_keynote_slide_transition_canonical_scaffold(root: Path) -> None:
    semantic = root / boundaries.KEYNOTE_SLIDE_TRANSITION_IMPLEMENTATION_SOURCES[0]
    semantic.parent.mkdir(parents=True, exist_ok=True)
    semantic_source = (
        semantic.read_text(encoding="utf-8") if semantic.is_file() else ""
    )
    semantic.write_text(
        semantic_source
        + "".join(
            f"pub struct {name};\n"
            for name in boundaries.KEYNOTE_SLIDE_TRANSITION_CANONICAL_TYPES
        ),
        encoding="utf-8",
    )
    lib_export = root / boundaries.KEYNOTE_SLIDE_TRANSITION_EXPORT_SOURCES[0]
    lib_export.parent.mkdir(parents=True, exist_ok=True)
    lib_source = (
        lib_export.read_text(encoding="utf-8") if lib_export.is_file() else ""
    )
    lib_export.write_text(lib_source + "pub mod transition;\n", encoding="utf-8")


def add_keynote_slide_delete_canonical_scaffold(root: Path) -> None:
    semantic = root / boundaries.KEYNOTE_SLIDE_DELETE_SEMANTIC_SOURCE
    semantic.parent.mkdir(parents=True, exist_ok=True)
    semantic.write_text(
        "".join(
            f"pub struct {name};\n"
            for name in boundaries.KEYNOTE_SLIDE_DELETE_CANONICAL_TYPES
        )
        + "impl Edit { pub fn remove_slide(&mut self, selector: SlideSelector) "
        "-> Result<&mut Self, Error> { todo!() } }\n",
        encoding="utf-8",
    )
    owner = root / boundaries.KEYNOTE_SLIDE_DELETE_OWNER_SOURCE
    owner.parent.mkdir(parents=True, exist_ok=True)
    owner.write_text(
        "impl Package {\n"
        "pub fn edit_slide_deletion(&self) -> slide::delete::Edit<'_> { todo!() }\n"
        "pub fn apply_slide_deletion(&self, patch: &slide::delete::Patch) "
        "-> Result<slide::delete::Commit, slide::delete::Error> { todo!() }\n"
        "}\n",
        encoding="utf-8",
    )
    lib_export, package_export, slide_export = (
        root / path for path in boundaries.KEYNOTE_SLIDE_DELETE_EXPORT_SOURCES
    )
    lib_export.write_text("pub mod slide;\n", encoding="utf-8")
    package_export.write_text("mod slide_delete;\n", encoding="utf-8")
    slide_export.write_text("pub mod delete;\n", encoding="utf-8")


def add_numbers_table_header_settings_canonical_scaffold(root: Path) -> None:
    semantic = root / boundaries.NUMBERS_TABLE_HEADER_SETTINGS_SEMANTIC_SOURCE
    semantic.parent.mkdir(parents=True, exist_ok=True)
    semantic_source = (
        semantic.read_text(encoding="utf-8") if semantic.is_file() else ""
    )
    semantic.write_text(
        semantic_source + "pub mod transaction;\n",
        encoding="utf-8",
    )
    transaction = root / boundaries.NUMBERS_TABLE_HEADER_SETTINGS_TRANSACTION_SOURCE
    transaction.parent.mkdir(parents=True, exist_ok=True)
    transaction_source = (
        transaction.read_text(encoding="utf-8") if transaction.is_file() else ""
    )
    transaction.write_text(
        transaction_source
        + "".join(
            f"pub struct {name};\n"
            for name in boundaries.NUMBERS_TABLE_HEADER_SETTINGS_CANONICAL_TYPES
        ),
        encoding="utf-8",
    )
    table_export = root / boundaries.NUMBERS_TABLE_HEADER_SETTINGS_EXPORT_SOURCES[2]
    table_export.parent.mkdir(parents=True, exist_ok=True)
    table_source = (
        table_export.read_text(encoding="utf-8") if table_export.is_file() else ""
    )
    table_export.write_text(table_source + "pub mod headers;\n", encoding="utf-8")
    lib_export = root / boundaries.NUMBERS_TABLE_HEADER_SETTINGS_EXPORT_SOURCES[0]
    lib_export.parent.mkdir(parents=True, exist_ok=True)
    lib_source = (
        lib_export.read_text(encoding="utf-8") if lib_export.is_file() else ""
    )
    lib_export.write_text(lib_source + "pub mod table;\n", encoding="utf-8")


def add_keynote_placeholder_visibility_canonical_scaffold(root: Path) -> None:
    semantic = root / boundaries.KEYNOTE_PLACEHOLDER_VISIBILITY_SEMANTIC_SOURCE
    semantic.parent.mkdir(parents=True, exist_ok=True)
    semantic_source = (
        semantic.read_text(encoding="utf-8") if semantic.is_file() else ""
    )
    semantic.write_text(
        semantic_source
        + "pub enum Kind { Title, Body, SlideNumber }\n"
        + "".join(
            f"pub struct {name};\n"
            for name in boundaries.KEYNOTE_PLACEHOLDER_VISIBILITY_CANONICAL_TYPES
            if name != "Kind"
        ),
        encoding="utf-8",
    )
    owner = root / boundaries.KEYNOTE_PLACEHOLDER_VISIBILITY_OWNER_SOURCE
    owner.parent.mkdir(parents=True, exist_ok=True)
    if not owner.is_file():
        owner.write_text("// Private package owner.\n", encoding="utf-8")
    lib_export = root / boundaries.KEYNOTE_PLACEHOLDER_VISIBILITY_EXPORT_SOURCES[0]
    lib_export.parent.mkdir(parents=True, exist_ok=True)
    lib_source = (
        lib_export.read_text(encoding="utf-8") if lib_export.is_file() else ""
    )
    lib_export.write_text(lib_source + "pub mod slide;\n", encoding="utf-8")
    package_export = root / boundaries.KEYNOTE_PLACEHOLDER_VISIBILITY_EXPORT_SOURCES[1]
    package_source = (
        package_export.read_text(encoding="utf-8")
        if package_export.is_file()
        else ""
    )
    package_export.write_text(
        package_source + "mod slide_placeholder_visibility;\n", encoding="utf-8"
    )
    slide_export = root / boundaries.KEYNOTE_PLACEHOLDER_VISIBILITY_EXPORT_SOURCES[2]
    slide_source = (
        slide_export.read_text(encoding="utf-8") if slide_export.is_file() else ""
    )
    slide_export.write_text(
        slide_source + "pub mod placeholder;\n", encoding="utf-8"
    )


def add_keynote_soundtrack_settings_canonical_scaffold(root: Path) -> None:
    semantic = root / boundaries.KEYNOTE_SOUNDTRACK_SETTINGS_SEMANTIC_SOURCE
    semantic.parent.mkdir(parents=True, exist_ok=True)
    semantic_source = (
        semantic.read_text(encoding="utf-8") if semantic.is_file() else ""
    )
    semantic.write_text(
        semantic_source
        + "".join(
            f"pub struct {name};\n"
            for name in boundaries.KEYNOTE_SOUNDTRACK_SETTINGS_CANONICAL_TYPES
        ),
        encoding="utf-8",
    )
    owner = root / boundaries.KEYNOTE_SOUNDTRACK_SETTINGS_OWNER_SOURCE
    owner.parent.mkdir(parents=True, exist_ok=True)
    if not owner.is_file():
        owner.write_text("// Private package owner.\n", encoding="utf-8")
    lib_export = root / boundaries.KEYNOTE_SOUNDTRACK_SETTINGS_EXPORT_SOURCES[0]
    lib_source = (
        lib_export.read_text(encoding="utf-8") if lib_export.is_file() else ""
    )
    lib_export.write_text(lib_source + "pub mod soundtrack;\n", encoding="utf-8")
    package_export = root / boundaries.KEYNOTE_SOUNDTRACK_SETTINGS_EXPORT_SOURCES[1]
    package_source = (
        package_export.read_text(encoding="utf-8")
        if package_export.is_file()
        else ""
    )
    package_export.write_text(
        package_source + "mod soundtrack_settings;\n", encoding="utf-8"
    )


def add_numbers_sheet_order_canonical_scaffold(root: Path) -> None:
    semantic = root / boundaries.NUMBERS_SHEET_ORDER_SEMANTIC_SOURCE
    semantic.parent.mkdir(parents=True, exist_ok=True)
    semantic_source = (
        semantic.read_text(encoding="utf-8") if semantic.is_file() else ""
    )
    semantic.write_text(
        semantic_source
        + "".join(
            f"pub struct {name};\n"
            for name in boundaries.NUMBERS_SHEET_ORDER_CANONICAL_TYPES
        ),
        encoding="utf-8",
    )
    owner = root / boundaries.NUMBERS_SHEET_ORDER_OWNER_SOURCE
    owner.parent.mkdir(parents=True, exist_ok=True)
    if not owner.is_file():
        owner.write_text("// Private package owner.\n", encoding="utf-8")
    lib_export = root / boundaries.NUMBERS_SHEET_ORDER_EXPORT_SOURCES[0]
    lib_source = (
        lib_export.read_text(encoding="utf-8") if lib_export.is_file() else ""
    )
    lib_export.write_text(lib_source + "pub mod sheet;\n", encoding="utf-8")
    package_export = root / boundaries.NUMBERS_SHEET_ORDER_EXPORT_SOURCES[1]
    package_source = (
        package_export.read_text(encoding="utf-8")
        if package_export.is_file()
        else ""
    )
    package_export.write_text(
        package_source + "mod sheet_order;\n", encoding="utf-8"
    )
    sheet_export = root / boundaries.NUMBERS_SHEET_ORDER_EXPORT_SOURCES[2]
    sheet_source = (
        sheet_export.read_text(encoding="utf-8") if sheet_export.is_file() else ""
    )
    sheet_export.write_text(sheet_source + "pub mod order;\n", encoding="utf-8")


def add_numbers_table_title_settings_canonical_scaffold(root: Path) -> None:
    semantic = root / boundaries.NUMBERS_TABLE_TITLE_SETTINGS_SEMANTIC_SOURCE
    semantic.parent.mkdir(parents=True, exist_ok=True)
    semantic.write_text(
        "".join(
            f"pub struct {name};\n"
            for name in boundaries.NUMBERS_TABLE_TITLE_SETTINGS_CANONICAL_TYPES
        ),
        encoding="utf-8",
    )
    owner = root / boundaries.NUMBERS_TABLE_TITLE_SETTINGS_OWNER_SOURCE
    owner.parent.mkdir(parents=True, exist_ok=True)
    owner.write_text("impl Package {}\n", encoding="utf-8")
    lib_export, package_export, table_export = (
        root / path for path in boundaries.NUMBERS_TABLE_TITLE_SETTINGS_EXPORT_SOURCES
    )
    lib_export.write_text("pub mod table;\n", encoding="utf-8")
    package_export.write_text("pub(crate) mod table_title;\n", encoding="utf-8")
    table_export.write_text("pub mod title;\n", encoding="utf-8")


def add_numbers_table_dimension_canonical_scaffold(root: Path) -> None:
    semantic = root / boundaries.NUMBERS_TABLE_DIMENSION_SEMANTIC_SOURCE
    semantic.parent.mkdir(parents=True, exist_ok=True)
    semantic.write_text(
        "pub mod transaction;\n"
        + "".join(
            f"pub struct {name};\n"
            for name in boundaries.NUMBERS_TABLE_DIMENSION_SEMANTIC_TYPES
        ),
        encoding="utf-8",
    )
    transaction = root / boundaries.NUMBERS_TABLE_DIMENSION_TRANSACTION_SOURCE
    transaction.parent.mkdir(parents=True, exist_ok=True)
    transaction.write_text(
        "".join(
            f"pub struct {name};\n"
            for name in boundaries.NUMBERS_TABLE_DIMENSION_TRANSACTION_TYPES
        ),
        encoding="utf-8",
    )
    owner = root / boundaries.NUMBERS_TABLE_DIMENSION_OWNER_SOURCE
    owner.parent.mkdir(parents=True, exist_ok=True)
    owner.write_text(
        "impl Package {\n"
        + "".join(
            f"pub fn {name}(&self, selector: TableSelector, dimension: Dimension) {{}}\n"
            for name in boundaries.NUMBERS_TABLE_DIMENSION_PACKAGE_METHODS
        )
        + "}\n",
        encoding="utf-8",
    )
    lib_export, package_export, table_export, _semantic_export = (
        root / path for path in boundaries.NUMBERS_TABLE_DIMENSION_EXPORT_SOURCES
    )
    lib_export.write_text(
        "pub mod table;\n"
        "pub use table::dimension::{Dimension, Points, Size};\n",
        encoding="utf-8",
    )
    package_export.write_text("pub(crate) mod table_dimension;\n", encoding="utf-8")
    table_export.write_text("pub mod dimension;\n", encoding="utf-8")


def add_numbers_table_cells_read_scaffold(root: Path) -> None:
    semantic = root / boundaries.NUMBERS_TABLE_CELLS_SEMANTIC_SOURCE
    semantic.parent.mkdir(parents=True, exist_ok=True)
    semantic.write_text(
        "pub struct State;\n"
        "pub enum Storage { Empty, Stored }\n"
        "pub struct Error;\n"
        "pub struct LimitKind;\n"
        "pub struct Path;\n",
        encoding="utf-8",
    )
    owner = root / boundaries.NUMBERS_TABLE_CELLS_OWNER_SOURCE
    owner.parent.mkdir(parents=True, exist_ok=True)
    owner.write_text(
        "impl Package {\n"
        "pub fn table_cell() {}\n"
        "pub fn table_cells() {}\n"
        "}\n",
        encoding="utf-8",
    )
    lib_export, package_export, table_export = (
        root / path for path in boundaries.NUMBERS_TABLE_CELLS_EXPORT_SOURCES
    )
    lib_export.write_text("pub mod table;\n", encoding="utf-8")
    package_export.write_text("pub(crate) mod table_cells;\n", encoding="utf-8")
    table_export.write_text("pub mod cells;\n", encoding="utf-8")


def add_numbers_table_cells_mutation_scaffold(root: Path) -> None:
    add_numbers_table_cells_read_scaffold(root)
    semantic = root / boundaries.NUMBERS_TABLE_CELLS_SEMANTIC_SOURCE
    semantic.write_text(
        "".join(
            f"pub struct {name};\n"
            for name in boundaries.NUMBERS_TABLE_CELLS_FULL_CANONICAL_TYPES
        ),
        encoding="utf-8",
    )
    owner = root / boundaries.NUMBERS_TABLE_CELLS_MUTATION_OWNER_SOURCE
    owner.write_text(
        "impl Package {\n"
        + "".join(
            f"pub fn {name}() {{}}\n"
            for name in boundaries.NUMBERS_TABLE_CELLS_MUTATION_PACKAGE_METHODS
        )
        + "}\n",
        encoding="utf-8",
    )
    package_export = root / boundaries.NUMBERS_TABLE_CELLS_EXPORT_SOURCES[1]
    package_export.write_text(
        package_export.read_text(encoding="utf-8")
        + "pub(crate) mod table_cell_edit;\n",
        encoding="utf-8",
    )


def add_pages_section_settings_canonical_scaffold(root: Path) -> None:
    semantic = root / boundaries.PAGES_SECTION_SETTINGS_SEMANTIC_SOURCE
    semantic.parent.mkdir(parents=True, exist_ok=True)
    semantic.write_text(
        "".join(
            f"pub struct {name};\n"
            for name in boundaries.PAGES_SECTION_SETTINGS_CANONICAL_TYPES
        ),
        encoding="utf-8",
    )
    owner = root / boundaries.PAGES_SECTION_SETTINGS_OWNER_SOURCE
    owner.parent.mkdir(parents=True, exist_ok=True)
    owner.write_text(
        "impl Package {\n"
        + "".join(
            f"pub fn {method}() {{}}\n"
            for method in boundaries.PAGES_SECTION_SETTINGS_PACKAGE_METHODS
        )
        + "}\n",
        encoding="utf-8",
    )
    lib_export, package_export, section_export = (
        root / path for path in boundaries.PAGES_SECTION_SETTINGS_EXPORT_SOURCES
    )
    lib_export.write_text("pub mod section;\n", encoding="utf-8")
    package_export.write_text("pub(crate) mod section_settings;\n", encoding="utf-8")
    section_export.write_text(
        "pub mod settings;\npub struct Settings;\n", encoding="utf-8"
    )


def add_pages_section_background_canonical_scaffold(root: Path) -> None:
    semantic = root / boundaries.PAGES_SECTION_BACKGROUND_SEMANTIC_SOURCE
    semantic.parent.mkdir(parents=True, exist_ok=True)
    semantic.write_text(
        "".join(
            f"pub struct {name};\n"
            for name in boundaries.PAGES_SECTION_BACKGROUND_CANONICAL_TYPES
        )
        + "impl Edit {\n"
        + "".join(
            f"pub fn {method}() {{}}\n"
            for method in boundaries.PAGES_SECTION_BACKGROUND_EDIT_METHODS
        )
        + "}\n",
        encoding="utf-8",
    )
    owner = root / boundaries.PAGES_SECTION_BACKGROUND_OWNER_SOURCE
    owner.parent.mkdir(parents=True, exist_ok=True)
    owner.write_text(
        "impl Package {\n"
        + "".join(
            f"pub fn {method}() {{}}\n"
            for method in boundaries.PAGES_SECTION_BACKGROUND_PACKAGE_METHODS
        )
        + "}\n",
        encoding="utf-8",
    )
    lib_export, package_export, section_export = (
        root / path for path in boundaries.PAGES_SECTION_BACKGROUND_EXPORT_SOURCES
    )
    lib_export.write_text("pub mod section;\n", encoding="utf-8")
    package_export.write_text(
        "pub(crate) mod section_background;\n", encoding="utf-8"
    )
    section_export.write_text(
        "pub mod background;\npub struct Background;\n", encoding="utf-8"
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
                    boundaries.Edge("litchi-odc", "soapberry-zip"),
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

    def test_numbers_metadata_core_edge_is_canonical(self) -> None:
        edge = boundaries.Edge("litchi-numbers", "litchi-core")

        self.assertIn(edge, self.policy.canonical_edges)
        self.assertNotIn(edge, self.policy.migration_edges)
        self.assertNotIn(edge, self.policy.dev_only_edges)

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

    def test_retired_iwa_keynote_document_reader_inventory_is_exact(self) -> None:
        self.assertEqual(
            boundaries.RETIRED_IWA_KEYNOTE_DOCUMENT_SOURCE,
            Path("crates/litchi-iwa/src/keynote/document.rs"),
        )
        self.assertEqual(
            boundaries.RETIRED_IWA_KEYNOTE_DOCUMENT_TYPES,
            ("KeynoteDocument", "KeynoteDocumentState", "KeynoteDocumentStats"),
        )
        self.assertEqual(
            boundaries.IWA_KEYNOTE_MODULE_SOURCE,
            Path("crates/litchi-iwa/src/keynote/mod.rs"),
        )
        self.assertEqual(
            boundaries.IWA_KEYNOTE_DOCUMENT_CALLER_ROOTS,
            (
                Path("crates/litchi-iwa/src"),
                Path("crates/litchi-iwa/tests"),
                Path("crates/litchi-iwa/examples"),
            ),
        )

    def test_retired_iwa_keynote_document_reader_surface_cannot_return(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            retired = root / boundaries.RETIRED_IWA_KEYNOTE_DOCUMENT_SOURCE
            retired.parent.mkdir(parents=True)
            retired.write_text("// retired reader returned\n", encoding="utf-8")
            module = root / boundaries.IWA_KEYNOTE_MODULE_SOURCE
            module.write_text(
                "pub mod r#document;\n"
                "pub use self::r#document::*;\n",
                encoding="utf-8",
            )
            source_caller = root / "crates/litchi-iwa/src/legacy_keynote.rs"
            source_caller.parent.mkdir(parents=True, exist_ok=True)
            source_caller.write_text(
                "pub fn open() -> KeynoteDocument { todo!() }\n"
                "pub type State = KeynoteDocumentState;\n"
                "/// Do not restore `KeynoteDocumentStats`.\n",
                encoding="utf-8",
            )
            test_caller = root / "crates/litchi-iwa/tests/keynote_reader.rs"
            test_caller.parent.mkdir(parents=True)
            test_caller.write_text(
                "fn assert_stats(_: KeynoteDocumentStats) {}\n",
                encoding="utf-8",
            )
            example_caller = root / "crates/litchi-iwa/examples/read_keynote.rs"
            example_caller.parent.mkdir(parents=True)
            example_caller.write_text(
                "use litchi_iwa::keynote::KeynoteDocument;\n",
                encoding="utf-8",
            )
            readme = root / boundaries.IWA_KEYNOTE_README
            readme.write_text(
                "Open with `KeynoteDocumentState` and inspect "
                "`KeynoteDocumentStats`.\n",
                encoding="utf-8",
            )

            self.assertEqual(
                boundaries.audit_iwa_keynote_document_source_topology(root),
                sorted(
                    [
                        "retired litchi-iwa Keynote document reader local re-export "
                        "document: crates/litchi-iwa/src/keynote/mod.rs:2",
                        "retired litchi-iwa Keynote document reader module document: "
                        "crates/litchi-iwa/src/keynote/mod.rs:1",
                        "retired litchi-iwa Keynote document reader README reference "
                        "KeynoteDocumentState: crates/litchi-iwa/README.md:1",
                        "retired litchi-iwa Keynote document reader README reference "
                        "KeynoteDocumentStats: crates/litchi-iwa/README.md:1",
                        "retired litchi-iwa Keynote document reader source returned: "
                        "crates/litchi-iwa/src/keynote/document.rs",
                        "retired litchi-iwa Keynote document reader rustdoc reference "
                        "KeynoteDocumentStats: crates/litchi-iwa/src/legacy_keynote.rs:3",
                        "retired litchi-iwa Keynote document reader type usage "
                        "KeynoteDocument: crates/litchi-iwa/examples/read_keynote.rs:1",
                        "retired litchi-iwa Keynote document reader type usage "
                        "KeynoteDocument: crates/litchi-iwa/src/legacy_keynote.rs:1",
                        "retired litchi-iwa Keynote document reader type usage "
                        "KeynoteDocumentState: crates/litchi-iwa/src/legacy_keynote.rs:2",
                        "retired litchi-iwa Keynote document reader type usage "
                        "KeynoteDocumentStats: crates/litchi-iwa/tests/keynote_reader.rs:1",
                    ]
                ),
            )

    def test_retired_iwa_keynote_document_rustdoc_references_cannot_return(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            path = root / boundaries.IWA_KEYNOTE_MODULE_SOURCE
            path.parent.mkdir(parents=True)
            path.write_text(
                "//! Use `KeynoteDocument`.\n"
                "/// State was `KeynoteDocumentState`.\n"
                "/** Stats were `KeynoteDocumentStats`. */\n"
                '#[doc = "Do not restore KeynoteDocument"]\n'
                "pub struct Reader;\n",
                encoding="utf-8",
            )

            self.assertEqual(
                boundaries.audit_iwa_keynote_document_source_topology(root),
                [
                    "retired litchi-iwa Keynote document reader rustdoc reference "
                    "KeynoteDocument: crates/litchi-iwa/src/keynote/mod.rs:1",
                    "retired litchi-iwa Keynote document reader rustdoc reference "
                    "KeynoteDocument: crates/litchi-iwa/src/keynote/mod.rs:4",
                    "retired litchi-iwa Keynote document reader rustdoc reference "
                    "KeynoteDocumentState: crates/litchi-iwa/src/keynote/mod.rs:2",
                    "retired litchi-iwa Keynote document reader rustdoc reference "
                    "KeynoteDocumentStats: crates/litchi-iwa/src/keynote/mod.rs:3",
                ],
            )

    def test_retired_iwa_keynote_document_module_and_reexport_variants(
        self,
    ) -> None:
        declarations = (
            ("mod document;", "module document"),
            ("pub(crate) mod r#document;", "module document"),
            ("pub\nmod\ndocument\n{}", "module document"),
            ("pub use document::*;", "local re-export document"),
            ("pub(crate) use self::r#document as legacy;", "local re-export document"),
        )
        for declaration, fragment in declarations:
            with self.subTest(declaration=declaration):
                with tempfile.TemporaryDirectory() as directory:
                    root = Path(directory)
                    module = root / boundaries.IWA_KEYNOTE_MODULE_SOURCE
                    module.parent.mkdir(parents=True)
                    module.write_text(declaration + "\n", encoding="utf-8")

                    violations = (
                        boundaries.audit_iwa_keynote_document_source_topology(root)
                    )
                    self.assertEqual(len(violations), 1)
                    self.assertIn(fragment, violations[0])

    def test_iwa_keynote_document_reader_policy_allows_builder_and_focused_reader(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            module = root / boundaries.IWA_KEYNOTE_MODULE_SOURCE
            module.parent.mkdir(parents=True)
            module.write_text(
                "pub use creation::KeynoteDocumentBuilder;\n"
                "pub use litchi_keynote::document::Document;\n"
                "pub use litchi_keynote::Package;\n",
                encoding="utf-8",
            )
            for relative in (
                Path("crates/litchi-iwa/src/keynote/creation.rs"),
                Path("crates/litchi-iwa/tests/generated_roundtrip.rs"),
                Path("crates/litchi-iwa/examples/create_keynote.rs"),
            ):
                path = root / relative
                path.parent.mkdir(parents=True, exist_ok=True)
                path.write_text(
                    "use litchi_iwa::keynote::KeynoteDocumentBuilder;\n"
                    "use litchi_keynote::{Package, document::Document};\n"
                    "pub struct KeynoteDocuments;\n"
                    "pub struct LegacyKeynoteDocumentStats;\n"
                    "// KeynoteDocument and KeynoteDocumentState are retired.\n",
                    encoding="utf-8",
                )
            readme = root / boundaries.IWA_KEYNOTE_README
            readme.write_text(
                "Use `KeynoteDocumentBuilder`, `litchi_keynote::Package`, or "
                "`litchi_keynote::document::Document`.\n",
                encoding="utf-8",
            )

            self.assertEqual(
                boundaries.audit_iwa_keynote_document_source_topology(root), []
            )

    def test_retired_iwa_keynote_show_settings_inventory_is_exact(self) -> None:
        self.assertEqual(
            boundaries.RETIRED_IWA_KEYNOTE_SHOW_SETTINGS_METHODS,
            ("show_settings", "set_show_settings"),
        )
        self.assertEqual(
            boundaries.RETIRED_IWA_KEYNOTE_SHOW_SETTINGS_MODULES,
            ("show_settings",),
        )
        self.assertEqual(
            boundaries.RETIRED_IWA_KEYNOTE_SHOW_SETTINGS_SOURCE,
            Path("crates/litchi-iwa/src/keynote/editor/show_settings.rs"),
        )
        self.assertEqual(
            boundaries.RETIRED_IWA_KEYNOTE_SHOW_SETTINGS_EXAMPLE,
            Path("crates/litchi-iwa/examples/edit_keynote_show.rs"),
        )
        self.assertEqual(
            boundaries.IWA_KEYNOTE_README,
            Path("crates/litchi-iwa/README.md"),
        )
        self.assertEqual(
            boundaries.KEYNOTE_SHOW_SETTINGS_FLAT_ALIASES,
            frozenset(
                {
                    "ShowSettings",
                    "ShowSettingsEdit",
                    "ShowSettingsPatch",
                    "ShowSettingsCommit",
                    "ShowSettingsDiagnostics",
                    "ShowSettingsError",
                    "ShowSettingsLimitKind",
                }
            ),
        )
        self.assertEqual(
            boundaries.KEYNOTE_SHOW_SETTINGS_SHORT_NAMES,
            frozenset(
                {
                    "Mode",
                    "Settings",
                    "Size",
                    "Show",
                    "Edit",
                    "Patch",
                    "Commit",
                    "Diagnostics",
                    "Error",
                    "LimitKind",
                }
            ),
        )
        self.assertEqual(
            boundaries.KEYNOTE_SHOW_SETTINGS_FLAT_SEMANTIC_ALIASES,
            frozenset({"Mode", "Settings", "Show", "Size"}),
        )

    def test_retired_iwa_keynote_show_settings_surface_cannot_return(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            nested = root / boundaries.IWA_KEYNOTE_SOURCE_ROOT / "legacy/settings.rs"
            nested.parent.mkdir(parents=True)
            nested.write_text(
                "fn r#show_settings() {}\n"
                "pub(in crate::keynote) async unsafe fn set_show_settings() {}\n",
                encoding="utf-8",
            )
            editor = root / boundaries.IWA_KEYNOTE_EDITOR_SOURCE
            editor.write_text("pub(crate) mod r#show_settings;\n", encoding="utf-8")
            retired = root / boundaries.RETIRED_IWA_KEYNOTE_SHOW_SETTINGS_SOURCE
            retired.parent.mkdir(parents=True, exist_ok=True)
            retired.write_text("// retired owner returned\n", encoding="utf-8")
            example = root / boundaries.RETIRED_IWA_KEYNOTE_SHOW_SETTINGS_EXAMPLE
            example.parent.mkdir(parents=True, exist_ok=True)
            example.write_text("// retired example returned\n", encoding="utf-8")

            self.assertEqual(
                boundaries.audit_iwa_keynote_show_settings_source_topology(root),
                sorted(
                    [
                        "retired litchi-iwa Keynote show-settings method "
                        "show_settings: "
                        "crates/litchi-iwa/src/keynote/legacy/settings.rs:1",
                        "retired litchi-iwa Keynote show-settings method "
                        "set_show_settings: "
                        "crates/litchi-iwa/src/keynote/legacy/settings.rs:2",
                        "retired litchi-iwa Keynote show-settings module show_settings: "
                        "crates/litchi-iwa/src/keynote/editor.rs:1",
                        "retired litchi-iwa Keynote show-settings source returned: "
                        "crates/litchi-iwa/src/keynote/editor/show_settings.rs",
                        "retired litchi-iwa Keynote show-settings example returned: "
                        "crates/litchi-iwa/examples/edit_keynote_show.rs",
                    ]
                ),
            )

    def test_retired_iwa_keynote_show_settings_module_declaration_variants(
        self,
    ) -> None:
        for declaration in (
            "mod show_settings;",
            "pub(crate) mod show_settings;",
            "pub(super) mod r#show_settings {}",
            "pub(in crate::keynote)\nmod\nr#show_settings\n{}",
            "pub mod show_settings { pub struct Legacy; }",
        ):
            with self.subTest(declaration=declaration):
                with tempfile.TemporaryDirectory() as directory:
                    root = Path(directory)
                    editor = root / boundaries.IWA_KEYNOTE_EDITOR_SOURCE
                    editor.parent.mkdir(parents=True)
                    editor.write_text(declaration + "\n", encoding="utf-8")

                    self.assertEqual(
                        boundaries.audit_iwa_keynote_show_settings_source_topology(
                            root
                        ),
                        [
                            "retired litchi-iwa Keynote show-settings module "
                            "show_settings: "
                            "crates/litchi-iwa/src/keynote/editor.rs:1"
                        ],
                    )

    def test_iwa_keynote_show_settings_policy_ignores_trivia_near_names_and_owners(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            host = root / boundaries.IWA_KEYNOTE_SOURCE_ROOT / "legacy/settings_old.rs"
            host.parent.mkdir(parents=True)
            host.write_text(
                "\n".join(
                    [
                        "// pub fn show_settings() {}",
                        'const NOTE: &str = "fn set_show_settings() {}";',
                        "/* fn show_settings() {}",
                        "   /* fn set_show_settings() {} */",
                        "   fn show_settings() {} */",
                        'const RAW_NOTE: &str = r###"fn set_show_settings() {}"###;',
                        "pub fn show_settings_snapshot() {}",
                        "pub fn reset_show_settings() {}",
                        "pub fn set_show_settings_for_show() {}",
                    ]
                )
                + "\n",
                encoding="utf-8",
            )
            editor = root / boundaries.IWA_KEYNOTE_EDITOR_SOURCE
            editor.write_text(
                "\n".join(
                    [
                        "// mod show_settings;",
                        'const NOTE: &str = "mod show_settings;";',
                        "/* mod show_settings;",
                        "   /* pub mod r#show_settings {} */",
                        "*/",
                        'const RAW_NOTE: &str = r#"mod show_settings;"#;',
                        "mod show_settings_legacy;",
                        "use crate::show_settings;",
                    ]
                )
                + "\n",
                encoding="utf-8",
            )
            declarations = (
                "pub fn show_settings() {}\npub fn set_show_settings() {}\n"
            )
            for relative in (
                Path("crates/litchi-keynote/src/package/show_settings.rs"),
                Path("crates/litchi-iwa/src/pages/editor/show_settings.rs"),
            ):
                path = root / relative
                path.parent.mkdir(parents=True, exist_ok=True)
                path.write_text(declarations, encoding="utf-8")
            non_rust = root / boundaries.IWA_KEYNOTE_SOURCE_ROOT / "show_settings.txt"
            non_rust.write_text(declarations, encoding="utf-8")

            self.assertEqual(
                boundaries.audit_iwa_keynote_show_settings_source_topology(root), []
            )

    def test_retired_iwa_keynote_show_settings_readme_calls_cannot_return(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            readme = root / boundaries.IWA_KEYNOTE_README
            readme.parent.mkdir(parents=True)
            readme.write_text(
                "\n".join(
                    [
                        "let current = keynote",
                        "    .",
                        "    r#show_settings",
                        "    (",
                        "let updated = package",
                        "    .r#set_show_settings",
                        "    (",
                        "let direct = crate::keynote::KeynoteEditor",
                        "    ::",
                        "    show_settings(",
                        "let direct = r#KeynoteEditor::r#set_show_settings(",
                        "let current = editor.show_settings(",
                    ]
                )
                + "\n",
                encoding="utf-8",
            )

            self.assertEqual(
                boundaries.audit_iwa_keynote_show_settings_source_topology(root),
                sorted(
                    [
                        "retired litchi-iwa Keynote show-settings README call "
                        "show_settings: crates/litchi-iwa/README.md:3",
                        "retired litchi-iwa Keynote show-settings README call "
                        "set_show_settings: crates/litchi-iwa/README.md:6",
                        "retired litchi-iwa Keynote show-settings README call "
                        "show_settings: crates/litchi-iwa/README.md:10",
                        "retired litchi-iwa Keynote show-settings README call "
                        "set_show_settings: crates/litchi-iwa/README.md:11",
                        "retired litchi-iwa Keynote show-settings README call "
                        "show_settings: crates/litchi-iwa/README.md:12",
                    ]
                ),
            )

    def test_iwa_keynote_show_settings_readme_call_policy_ignores_safe_text(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            readme = root / boundaries.IWA_KEYNOTE_README
            readme.parent.mkdir(parents=True)
            safe_text = "\n".join(
                [
                    "The show_settings and set_show_settings names are retired.",
                    "Use `show_settings` rather than `.show_settings`.",
                    "show_settings(",
                    "set_show_settings(",
                    "editor.show_settings_snapshot(",
                    "editor.set_show_settings_for_show(",
                    "KeynoteEditor::show_settings_snapshot(",
                    "Package::show_settings(",
                    "Package::set_show_settings(",
                    "package.show_settings(",
                    "show.show_settings(",
                ]
            ) + "\n"
            readme.write_text(safe_text, encoding="utf-8")
            for relative in (
                Path("README.md"),
                Path("crates/litchi-keynote/README.md"),
                Path("crates/litchi-iwa/examples/README.md"),
            ):
                other = root / relative
                other.parent.mkdir(parents=True, exist_ok=True)
                other.write_text(
                    "editor.show_settings(\neditor.set_show_settings(\n",
                    encoding="utf-8",
                )

            self.assertEqual(
                boundaries.audit_iwa_keynote_show_settings_source_topology(root), []
            )

    def test_focused_keynote_show_settings_public_api_rejects_physical_leaks(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            semantic, transaction = (
                root / path
                for path in boundaries.KEYNOTE_SHOW_SETTINGS_IMPLEMENTATION_SOURCES
            )
            semantic.parent.mkdir(parents=True)
            semantic.write_text(
                "\n".join(
                    [
                        "pub fn show_settings(r#source_bytes: &[u8], r#object_id: u64) "
                        "-> ShowArchive {}",
                        "pub type Patch = buffa::ShowArchiveView;",
                        "impl prost::Message for show::Edit {}",
                    ]
                )
                + "\n",
                encoding="utf-8",
            )
            transaction.parent.mkdir(parents=True, exist_ok=True)
            transaction.write_text(
                "pub type Commit = IWorkPackage;\n"
                "pub type Diagnostics = litchi_iwa_core::RawObject;\n"
                "pub type Error = SourceBytes;\n"
                "pub type Settings = SourceCatalog;\n",
                encoding="utf-8",
            )
            lib_export, package_export = (
                root / path
                for path in boundaries.KEYNOTE_SHOW_SETTINGS_EXPORT_SOURCES
            )
            lib_export.write_text(
                "pub use litchi_iwa_protos::show::GeneratedArchive as Patch;\n"
                "pub type Patch = crate::show::GeneratedSettings;\n",
                encoding="utf-8",
            )
            package_export.write_text(
                "pub type LimitKind = "
                "crate::show::LimitKind<prost_types::MessageInfo>;\n",
                encoding="utf-8",
            )

            violations = (
                boundaries.audit_keynote_show_settings_facade_source_topology(root)
            )

            self.assertEqual(violations, sorted(violations))
            self.assertTrue(
                all(
                    violation.startswith(
                        "focused litchi-keynote show-settings public API exposes "
                    )
                    for violation in violations
                )
            )
            expected_fragments = (
                "raw source bytes source_bytes",
                "raw byte slice &[u8]",
                "raw identifier object_id",
                "archive/IWA type ShowArchive",
                "protobuf type buffa",
                "archive/IWA type ShowArchiveView",
                "protobuf type prost",
                "protobuf type Message",
                "archive/IWA type IWorkPackage",
                "archive/IWA type litchi_iwa_core",
                "native object RawObject",
                "raw source bytes SourceBytes",
                "archive/IWA type SourceCatalog",
                "archive/IWA type litchi_iwa_protos",
                "archive/IWA type GeneratedArchive",
                "generated type GeneratedSettings",
                "protobuf type prost_types",
                "archive/IWA type MessageInfo",
            )
            self.assertEqual(len(violations), len(expected_fragments))
            for fragment in expected_fragments:
                with self.subTest(fragment=fragment):
                    self.assertTrue(
                        any(fragment in violation for violation in violations),
                        msg=f"missing focused show-settings leak: {fragment}",
                    )

    def test_focused_keynote_show_settings_api_ignores_nested_and_private_code(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            semantic, transaction = (
                root / path
                for path in boundaries.KEYNOTE_SHOW_SETTINGS_IMPLEMENTATION_SOURCES
            )
            semantic.parent.mkdir(parents=True)
            semantic.write_text(
                "\n".join(
                    [
                        "// pub type ShowSettings = ShowArchive;",
                        'const NOTE: &str = "pub struct ShowSettingsEdit;";',
                        "/* pub type ShowSettingsPatch = buffa::ShowArchiveView; */",
                        "pub struct Mode;",
                        "pub struct Settings;",
                        "pub struct Size;",
                        "pub struct Show;",
                        "pub struct Edit;",
                        "pub struct Patch;",
                        "pub struct Commit;",
                        "pub struct Diagnostics;",
                        "pub struct Error;",
                        "pub struct LimitKind;",
                        "pub fn show(settings: Settings, input: InputBytes, "
                        "byte_count: usize) -> OutputBytes { todo!() }",
                        "fn source_bytes(source_bytes: &[u8], object_id: u64) "
                        "-> SourceBytes { todo!() }",
                        "pub(crate) fn restricted(archive: ShowArchive) {}",
                        "struct ShowSettings;",
                        "pub(crate) struct ShowSettingsEdit;",
                        "pub(super) type ShowSettingsPatch = Patch;",
                        "pub(in crate) struct ShowSettingsCommit;",
                        "impl ShowSettingsEdit {}",
                        "impl prost::Message for Unrelated {}",
                    ]
                )
                + "\n",
                encoding="utf-8",
            )
            transaction.parent.mkdir(parents=True, exist_ok=True)
            transaction.write_text(
                "struct ShowSettingsDiagnostics;\n"
                "pub(crate) struct ShowSettingsError;\n"
                "pub(super) type ShowSettingsLimitKind = LimitKind;\n"
                "impl ShowSettingsPatch {}\n",
                encoding="utf-8",
            )
            lib_export, package_export = (
                root / path
                for path in boundaries.KEYNOTE_SHOW_SETTINGS_EXPORT_SOURCES
            )
            canonical_exports = (
                "pub mod show;\n"
                "// pub use crate::show::{Mode, Settings};\n"
                'const SHOW_NOTE: &str = "pub use crate::show::{Show, Size};";\n'
                "pub fn show_settings(settings: show::Settings, mode: show::Mode, "
                "size: show::Size) -> show::Show { todo!() }\n"
                "use crate::show::Mode;\n"
                "pub(crate) use crate::show::Settings;\n"
                "pub(super) type Show = crate::show::Show;\n"
                "pub(in crate) use crate::show::Size;\n"
                "pub(crate) use crate::show::*;\n"
                "pub use crate::render::{Mode, Settings};\n"
                "pub use crate::render::*;\n"
                "pub type Show = crate::presentation::Show;\n"
                "pub use crate::geometry::Size;\n"
                "pub fn unrelated(object_id: u64) -> ShowArchive { todo!() }\n"
                "pub struct ShowSetting;\n"
                "pub struct ShowSettingsSnapshot;\n"
            )
            lib_export.write_text(canonical_exports, encoding="utf-8")
            package_export.write_text(canonical_exports, encoding="utf-8")
            other_owner = root / "crates/litchi-pages/src/show.rs"
            other_owner.parent.mkdir(parents=True)
            other_owner.write_text(
                "pub struct ShowSettings;\n"
                "pub type ShowSettingsPatch = litchi_iwa_core::RawObject;\n"
                "pub use crate::show::{Mode, Settings, Show, Size};\n",
                encoding="utf-8",
            )

            self.assertEqual(
                boundaries.audit_keynote_show_settings_facade_source_topology(root),
                [],
            )

    def test_focused_keynote_show_settings_exports_reject_flat_semantic_aliases(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            lib_export, package_export = (
                root / path
                for path in boundaries.KEYNOTE_SHOW_SETTINGS_EXPORT_SOURCES
            )
            lib_export.parent.mkdir(parents=True)
            lib_export.write_text(
                "pub use crate::show::{Mode, Settings as Settings};\n"
                "pub type Show = crate::show::Show;\n",
                encoding="utf-8",
            )
            package_export.write_text(
                "pub use crate::show_settings::Size;\n"
                "pub use crate::show::*;\n",
                encoding="utf-8",
            )

            self.assertEqual(
                boundaries.audit_keynote_show_settings_facade_source_topology(root),
                sorted(
                    [
                        "focused litchi-keynote show-settings public API retains "
                        "flat semantic alias Mode: "
                        "crates/litchi-keynote/src/lib.rs:1",
                        "focused litchi-keynote show-settings public API retains "
                        "flat semantic alias Settings: "
                        "crates/litchi-keynote/src/lib.rs:1",
                        "focused litchi-keynote show-settings public API retains "
                        "flat semantic alias Show: "
                        "crates/litchi-keynote/src/lib.rs:2",
                        "focused litchi-keynote show-settings public API retains "
                        "flat semantic alias Size: "
                        "crates/litchi-keynote/src/package.rs:1",
                        "focused litchi-keynote show-settings public API retains "
                        "flat semantic aliases via show glob: "
                        "crates/litchi-keynote/src/package.rs:2",
                    ]
                ),
            )

    def test_focused_keynote_show_settings_public_api_rejects_flat_aliases(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            semantic, transaction = (
                root / path
                for path in boundaries.KEYNOTE_SHOW_SETTINGS_IMPLEMENTATION_SOURCES
            )
            semantic.parent.mkdir(parents=True)
            semantic.write_text(
                "pub struct ShowSettings;\n"
                "pub struct ShowSettingsEdit;\n"
                "pub type ShowSettingsPatch = Patch;\n",
                encoding="utf-8",
            )
            transaction.parent.mkdir(parents=True, exist_ok=True)
            transaction.write_text(
                "pub type ShowSettingsCommit = Commit;\n"
                "pub type ShowSettingsDiagnostics = Diagnostics;\n",
                encoding="utf-8",
            )
            lib_export, package_export = (
                root / path
                for path in boundaries.KEYNOTE_SHOW_SETTINGS_EXPORT_SOURCES
            )
            lib_export.write_text(
                "pub use crate::show::Error as ShowSettingsError;\n",
                encoding="utf-8",
            )
            package_export.write_text(
                "pub use crate::show::LimitKind as ShowSettingsLimitKind;\n",
                encoding="utf-8",
            )

            self.assertEqual(
                boundaries.audit_keynote_show_settings_facade_source_topology(root),
                sorted(
                    [
                        "focused litchi-keynote show-settings public API retains "
                        "flat alias ShowSettings: "
                        "crates/litchi-keynote/src/show.rs:1",
                        "focused litchi-keynote show-settings public API retains "
                        "flat alias ShowSettingsEdit: "
                        "crates/litchi-keynote/src/show.rs:2",
                        "focused litchi-keynote show-settings public API retains "
                        "flat alias ShowSettingsPatch: "
                        "crates/litchi-keynote/src/show.rs:3",
                        "focused litchi-keynote show-settings public API retains "
                        "flat alias ShowSettingsCommit: "
                        "crates/litchi-keynote/src/package/show_settings.rs:1",
                        "focused litchi-keynote show-settings public API retains "
                        "flat alias ShowSettingsDiagnostics: "
                        "crates/litchi-keynote/src/package/show_settings.rs:2",
                        "focused litchi-keynote show-settings public API retains "
                        "flat alias ShowSettingsError: "
                        "crates/litchi-keynote/src/lib.rs:1",
                        "focused litchi-keynote show-settings public API retains "
                        "flat alias ShowSettingsLimitKind: "
                        "crates/litchi-keynote/src/package.rs:1",
                    ]
                ),
            )

    def test_keynote_soundtrack_settings_boundary_inventories_are_exact(self) -> None:
        self.assertEqual(
            boundaries.RETIRED_IWA_KEYNOTE_SOUNDTRACK_SETTINGS_METHODS,
            ("soundtrack_settings", "set_soundtrack_settings"),
        )
        self.assertEqual(
            boundaries.RETIRED_IWA_KEYNOTE_SOUNDTRACK_SETTINGS_SOURCE,
            Path("crates/litchi-iwa/src/keynote/editor/soundtrack.rs"),
        )
        self.assertEqual(
            boundaries.RETIRED_IWA_KEYNOTE_SOUNDTRACK_SETTINGS_MODULES,
            ("soundtrack",),
        )
        self.assertEqual(
            boundaries.RETIRED_IWA_KEYNOTE_SOUNDTRACK_SETTINGS_EXAMPLE,
            Path("crates/litchi-iwa/examples/edit_keynote_soundtrack.rs"),
        )
        self.assertEqual(
            boundaries.RETIRED_IWA_KEYNOTE_SOUNDTRACK_SETTINGS_TESTS,
            (
                "soundtrack_settings_are_typed_transactional_and_wire_exact",
                "soundtrack_settings_handle_absent_and_malformed_objects_transactionally",
            ),
        )
        self.assertEqual(
            boundaries.RETIRED_IWA_KEYNOTE_SOUNDTRACK_SETTINGS_WIRE_METHODS,
            ("patch_soundtrack_wire",),
        )
        self.assertEqual(
            boundaries.IWA_KEYNOTE_SOUNDTRACK_WIRE_SOURCE,
            Path("crates/litchi-iwa/src/keynote/editor/soundtrack_wire.rs"),
        )
        self.assertEqual(
            boundaries.IWA_KEYNOTE_SOUNDTRACK_SETTINGS_CALLER_SOURCES,
            (Path("crates/litchi-iwa/examples/inspect_keynote_structure.rs"),),
        )
        self.assertEqual(
            boundaries.KEYNOTE_SOUNDTRACK_SETTINGS_SEMANTIC_SOURCE,
            Path("crates/litchi-keynote/src/soundtrack.rs"),
        )
        self.assertEqual(
            boundaries.KEYNOTE_SOUNDTRACK_SETTINGS_OWNER_SOURCE,
            Path("crates/litchi-keynote/src/package/soundtrack_settings.rs"),
        )
        self.assertEqual(
            boundaries.KEYNOTE_SOUNDTRACK_SETTINGS_OWNER_HELPER_ROOT,
            Path("crates/litchi-keynote/src/package/soundtrack_settings"),
        )
        self.assertEqual(
            boundaries.KEYNOTE_SOUNDTRACK_SETTINGS_OWNER_HELPER_SOURCES,
            (
                Path(
                    "crates/litchi-keynote/src/package/soundtrack_settings/media.rs"
                ),
                Path(
                    "crates/litchi-keynote/src/package/soundtrack_settings/rewrite.rs"
                ),
            ),
        )
        self.assertEqual(
            boundaries.KEYNOTE_SOUNDTRACK_SETTINGS_IMPLEMENTATION_SOURCES,
            (
                Path("crates/litchi-keynote/src/soundtrack.rs"),
                Path("crates/litchi-keynote/src/package/soundtrack_settings.rs"),
                Path(
                    "crates/litchi-keynote/src/package/soundtrack_settings/media.rs"
                ),
                Path(
                    "crates/litchi-keynote/src/package/soundtrack_settings/rewrite.rs"
                ),
            ),
        )
        self.assertEqual(
            boundaries.KEYNOTE_SOUNDTRACK_SETTINGS_EXPORT_SOURCES,
            (
                Path("crates/litchi-keynote/src/lib.rs"),
                Path("crates/litchi-keynote/src/package.rs"),
            ),
        )
        self.assertEqual(
            boundaries.KEYNOTE_SOUNDTRACK_SETTINGS_CANONICAL_TYPES,
            ("Mode", "Settings", "Edit", "Patch", "Commit", "Diagnostics", "Error", "LimitKind"),
        )
        self.assertEqual(
            boundaries.KEYNOTE_SOUNDTRACK_SETTINGS_PACKAGE_METHODS,
            (
                "soundtrack_settings",
                "edit_soundtrack_settings",
                "apply_soundtrack_settings",
            ),
        )
        self.assertEqual(
            boundaries.KEYNOTE_SOUNDTRACK_SETTINGS_FLAT_ALIASES,
            frozenset(
                {
                    "Soundtrack",
                    "SoundtrackMode",
                    "SoundtrackSettings",
                    "SoundtrackEdit",
                    "SoundtrackPatch",
                    "SoundtrackCommit",
                    "SoundtrackDiagnostics",
                    "SoundtrackError",
                    "SoundtrackLimitKind",
                    "SoundtrackSettingsEdit",
                    "SoundtrackSettingsPatch",
                    "SoundtrackSettingsCommit",
                    "SoundtrackSettingsDiagnostics",
                    "SoundtrackSettingsError",
                    "SoundtrackSettingsLimitKind",
                }
            ),
        )
        self.assertEqual(
            boundaries.KEYNOTE_SOUNDTRACK_SETTINGS_FORBIDDEN_PUBLIC_MEMBERS,
            frozenset({"set_soundtrack_settings"}),
        )
        self.assertEqual(
            boundaries.KEYNOTE_SOUNDTRACK_SETTINGS_PHYSICAL_TYPES,
            frozenset(
                {
                    "Archive",
                    "ArchiveObject",
                    "ComponentCatalog",
                    "EntryEdit",
                    "ExactArtifacts",
                    "IWorkPackage",
                    "PhysicalSource",
                    "RawMessage",
                    "ReferenceSnapshot",
                    "Resolved",
                    "SnappyStream",
                    "SoundtrackRecord",
                    "SoundtrackSnapshot",
                    "SoundtrackSettingsSnapshot",
                    "SourceCatalog",
                }
            ),
        )
        self.assertEqual(
            boundaries.KEYNOTE_SOUNDTRACK_SETTINGS_WIRE_TYPES,
            frozenset(
                {
                    "DecodeLimitKind",
                    "DecodeOptions",
                    "NestedFieldEdit",
                    "NestedFieldReplacement",
                    "WireDescent",
                    "WireError",
                    "WireLimits",
                    "WireResourceLimit",
                    "WireView",
                }
            ),
        )
        self.assertEqual(
            boundaries.KEYNOTE_SOUNDTRACK_SETTINGS_PROTO_ORIGINS,
            frozenset({"kn", "tsp"}),
        )
        self.assertEqual(
            boundaries.KEYNOTE_SOUNDTRACK_SETTINGS_MEDIA_TOPOLOGY_NAMES,
            frozenset(
                {
                    "DataReference",
                    "EmbeddedMediaAsset",
                    "KeynoteSoundtrackItemInfo",
                    "MediaAssetId",
                    "data_reference",
                    "data_references",
                    "media",
                    "media_items",
                    "movie_media",
                    "payload",
                    "payloads",
                }
            ),
        )

    def test_retired_iwa_keynote_soundtrack_settings_surface_cannot_return(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            retired = root / boundaries.RETIRED_IWA_KEYNOTE_SOUNDTRACK_SETTINGS_SOURCE
            retired.parent.mkdir(parents=True)
            retired.write_text(
                "pub fn r#soundtrack_settings() {}\n"
                "pub fn set_soundtrack_settings() {}\n",
                encoding="utf-8",
            )
            wire = root / boundaries.IWA_KEYNOTE_SOUNDTRACK_WIRE_SOURCE
            wire.write_text("fn r#patch_soundtrack_wire() {}\n", encoding="utf-8")
            editor = root / boundaries.IWA_KEYNOTE_EDITOR_SOURCE
            editor.write_text("pub mod r#soundtrack;\n", encoding="utf-8")
            tests = root / boundaries.IWA_KEYNOTE_EDITOR_TEST_SOURCE
            tests.write_text(
                "fn soundtrack_settings_are_typed_transactional_and_wire_exact() {}\n"
                "fn soundtrack_settings_handle_absent_and_malformed_objects_transactionally() {}\n",
                encoding="utf-8",
            )
            example = root / boundaries.RETIRED_IWA_KEYNOTE_SOUNDTRACK_SETTINGS_EXAMPLE
            example.parent.mkdir(parents=True)
            example.write_text("fn main() {}\n", encoding="utf-8")

            self.assertEqual(
                boundaries.audit_iwa_keynote_soundtrack_settings_source_topology(root),
                sorted(
                    [
                        "retired litchi-iwa Keynote soundtrack settings source returned: "
                        "crates/litchi-iwa/src/keynote/editor/soundtrack.rs",
                        "retired litchi-iwa Keynote soundtrack settings example returned: "
                        "crates/litchi-iwa/examples/edit_keynote_soundtrack.rs",
                        "retired litchi-iwa Keynote soundtrack settings method "
                        "soundtrack_settings: crates/litchi-iwa/src/keynote/editor/"
                        "soundtrack.rs:1",
                        "retired litchi-iwa Keynote soundtrack settings method "
                        "set_soundtrack_settings: crates/litchi-iwa/src/keynote/editor/"
                        "soundtrack.rs:2",
                        "retired litchi-iwa Keynote soundtrack settings wire helper "
                        "patch_soundtrack_wire: crates/litchi-iwa/src/keynote/editor/"
                        "soundtrack_wire.rs:1",
                        "retired litchi-iwa Keynote soundtrack settings module soundtrack: "
                        "crates/litchi-iwa/src/keynote/editor.rs:1",
                        "retired litchi-iwa Keynote soundtrack settings test "
                        "soundtrack_settings_are_typed_transactional_and_wire_exact: "
                        "crates/litchi-iwa/src/keynote/editor/tests.rs:1",
                        "retired litchi-iwa Keynote soundtrack settings test "
                        "soundtrack_settings_handle_absent_and_malformed_objects_"
                        "transactionally: crates/litchi-iwa/src/keynote/editor/tests.rs:2",
                    ]
                ),
            )

    def test_retired_iwa_keynote_soundtrack_settings_module_variants(self) -> None:
        for declaration in (
            "mod soundtrack;",
            "pub mod r#soundtrack;",
            "pub(crate) mod soundtrack {}",
            "pub(super) mod r#soundtrack {}",
            "pub(in crate) mod soundtrack;",
            "pub\nmod\nr#soundtrack\n{}",
        ):
            with self.subTest(declaration=declaration):
                with tempfile.TemporaryDirectory() as directory:
                    root = Path(directory)
                    editor = root / boundaries.IWA_KEYNOTE_EDITOR_SOURCE
                    editor.parent.mkdir(parents=True)
                    editor.write_text(declaration + "\n", encoding="utf-8")
                    expected_line = 3 if declaration.startswith("pub\n") else 1
                    self.assertEqual(
                        boundaries.audit_iwa_keynote_soundtrack_settings_source_topology(
                            root
                        ),
                        [
                            "retired litchi-iwa Keynote soundtrack settings module "
                            "soundtrack: crates/litchi-iwa/src/keynote/editor.rs:"
                            f"{expected_line}"
                        ],
                    )

    def test_retired_iwa_keynote_soundtrack_settings_calls_and_readme_example(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            caller_relative = boundaries.IWA_KEYNOTE_SOUNDTRACK_SETTINGS_CALLER_SOURCES[0]
            caller = root / caller_relative
            caller.parent.mkdir(parents=True)
            caller.write_text(
                "keynote.r#soundtrack_settings();\n"
                "canvas\n    .\n    set_soundtrack_settings(settings);\n"
                "crate::nested::KeynoteEditor::\n"
                "    r#soundtrack_settings();\n",
                encoding="utf-8",
            )
            readme = root / boundaries.IWA_KEYNOTE_README
            readme.parent.mkdir(parents=True, exist_ok=True)
            readme.write_text(
                "editor.soundtrack_settings();\n"
                "package\n  .set_soundtrack_settings(settings);\n"
                "edit_keynote_soundtrack\n"
                "edit_keynote_soundtrack.rs\n",
                encoding="utf-8",
            )

            self.assertEqual(
                boundaries.audit_iwa_keynote_soundtrack_settings_source_topology(root),
                sorted(
                    [
                        "retired litchi-iwa Keynote soundtrack settings caller "
                        f"soundtrack_settings: {caller_relative}:1",
                        "retired litchi-iwa Keynote soundtrack settings caller "
                        f"set_soundtrack_settings: {caller_relative}:4",
                        "retired litchi-iwa Keynote soundtrack settings caller "
                        f"soundtrack_settings: {caller_relative}:6",
                        "retired litchi-iwa Keynote soundtrack settings README call "
                        "soundtrack_settings: crates/litchi-iwa/README.md:1",
                        "retired litchi-iwa Keynote soundtrack settings README call "
                        "set_soundtrack_settings: crates/litchi-iwa/README.md:3",
                        "retired litchi-iwa Keynote soundtrack settings README example "
                        "reference edit_keynote_soundtrack: crates/litchi-iwa/README.md:4",
                        "retired litchi-iwa Keynote soundtrack settings README example "
                        "reference edit_keynote_soundtrack: crates/litchi-iwa/README.md:5",
                    ]
                ),
            )

    def test_iwa_keynote_soundtrack_settings_policy_retains_media_item_surfaces(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            editor = root / boundaries.IWA_KEYNOTE_EDITOR_SOURCE
            editor.parent.mkdir(parents=True)
            editor.write_text(
                "// mod soundtrack;\n"
                'const NOTE: &str = "pub mod soundtrack;";\n'
                'const RAW: &str = r#"mod r#soundtrack {}"#;\n'
                "/* outer /* mod soundtrack; */ still comment */\n"
                "mod soundtrack_wire;\n"
                "mod soundtrack_items;\n"
                "pub struct KeynoteSoundtrackItemInfo;\n",
                encoding="utf-8",
            )
            wire = root / boundaries.IWA_KEYNOTE_SOUNDTRACK_WIRE_SOURCE
            wire.parent.mkdir(parents=True, exist_ok=True)
            wire.write_text(
                "// fn patch_soundtrack_wire() {}\n"
                'const NOTE: &str = "fn patch_soundtrack_wire() {}";\n'
                "fn decode_soundtrack() {}\n"
                "fn patch_soundtrack_media_wire() {}\n"
                "fn replace_soundtrack_media() {}\n",
                encoding="utf-8",
            )
            items = root / boundaries.IWA_KEYNOTE_SOURCE_ROOT / "editor/soundtrack_items.rs"
            items.write_text(
                "pub fn soundtrack_items() {}\n"
                "pub fn add_soundtrack_item() {}\n"
                "pub fn insert_soundtrack_item() {}\n"
                "pub fn replace_soundtrack_item() {}\n"
                "pub fn move_soundtrack_item() {}\n"
                "pub fn remove_soundtrack_item() {}\n",
                encoding="utf-8",
            )
            tests = root / boundaries.IWA_KEYNOTE_EDITOR_TEST_SOURCE
            tests.write_text(
                "// fn soundtrack_settings_are_typed_transactional_and_wire_exact() {}\n"
                'const TEST_NOTE: &str = "fn soundtrack_settings_handle_absent_and_malformed_objects_transactionally() {}";\n'
                "fn soundtrack_items_round_trip_exactly() {}\n",
                encoding="utf-8",
            )
            caller = root / boundaries.IWA_KEYNOTE_SOUNDTRACK_SETTINGS_CALLER_SOURCES[0]
            caller.parent.mkdir(parents=True)
            caller.write_text(
                "// keynote.soundtrack_settings();\n"
                'const NOTE: &str = "editor.set_soundtrack_settings(settings);";\n'
                "package.soundtrack_settings();\n"
                "package.edit_soundtrack_settings();\n"
                "package.apply_soundtrack_settings(patch);\n",
                encoding="utf-8",
            )
            retained_example = (
                root / "crates/litchi-iwa/examples/edit_keynote_soundtrack_items.rs"
            )
            retained_example.write_text("fn main() {}\n", encoding="utf-8")
            other_owner = root / "crates/litchi-pages/src/editor.rs"
            other_owner.parent.mkdir(parents=True)
            other_owner.write_text(
                "pub fn soundtrack_settings() {}\n"
                "pub fn set_soundtrack_settings() {}\n",
                encoding="utf-8",
            )
            readme = root / boundaries.IWA_KEYNOTE_README
            readme.parent.mkdir(parents=True, exist_ok=True)
            readme.write_text(
                "`soundtrack_settings` is retired prose.\n"
                "set_soundtrack_settings\n"
                "package.soundtrack_settings()\n"
                "package.edit_soundtrack_settings()\n"
                "package.apply_soundtrack_settings(patch)\n"
                "edit_keynote_soundtrack_items.rs\n",
                encoding="utf-8",
            )

            self.assertEqual(
                boundaries.audit_iwa_keynote_soundtrack_settings_source_topology(root),
                [],
            )

    def test_focused_keynote_soundtrack_settings_requires_each_canonical_type(
        self,
    ) -> None:
        for missing in boundaries.KEYNOTE_SOUNDTRACK_SETTINGS_CANONICAL_TYPES:
            with self.subTest(missing=missing):
                with tempfile.TemporaryDirectory() as directory:
                    root = Path(directory)
                    add_keynote_soundtrack_settings_canonical_scaffold(root)
                    semantic = root / boundaries.KEYNOTE_SOUNDTRACK_SETTINGS_SEMANTIC_SOURCE
                    semantic.write_text(
                        "".join(
                            f"pub struct {name};\n"
                            for name in boundaries.KEYNOTE_SOUNDTRACK_SETTINGS_CANONICAL_TYPES
                            if name != missing
                        ),
                        encoding="utf-8",
                    )
                    self.assertEqual(
                        boundaries.audit_keynote_soundtrack_settings_facade_source_topology(
                            root
                        ),
                        [
                            "focused litchi-keynote soundtrack settings public API "
                            f"is missing canonical soundtrack type {missing}: "
                            "crates/litchi-keynote/src/soundtrack.rs"
                        ],
                    )

    def test_focused_keynote_soundtrack_settings_requires_root_and_private_owner(
        self,
    ) -> None:
        rejected_root = (
            "",
            "mod soundtrack;\n",
            "pub(crate) mod soundtrack;\n",
            "pub(super) mod r#soundtrack {}\n",
            "pub(in crate) mod soundtrack;\n",
            "// pub mod soundtrack;\n",
            'const NOTE: &str = "pub mod soundtrack;";\n',
        )
        accepted_root = (
            "pub mod soundtrack;\n",
            "pub mod r#soundtrack;\n",
            "pub mod soundtrack {}\n",
            "pub\nmod\nr#soundtrack\n{}\n",
        )
        for declaration, expected in (
            *[
                (
                    declaration,
                    [
                        "focused litchi-keynote soundtrack settings public API is "
                        "missing canonical root soundtrack module: "
                        "crates/litchi-keynote/src/lib.rs"
                    ],
                )
                for declaration in rejected_root
            ],
            *[(declaration, []) for declaration in accepted_root],
        ):
            with self.subTest(scope="root", declaration=declaration):
                with tempfile.TemporaryDirectory() as directory:
                    root = Path(directory)
                    add_keynote_soundtrack_settings_canonical_scaffold(root)
                    lib_export = root / boundaries.KEYNOTE_SOUNDTRACK_SETTINGS_EXPORT_SOURCES[0]
                    lib_export.write_text(declaration, encoding="utf-8")
                    self.assertEqual(
                        boundaries.audit_keynote_soundtrack_settings_facade_source_topology(
                            root
                        ),
                        expected,
                    )

        rejected_owner = (
            "",
            "// mod soundtrack_settings;\n",
            'const NOTE: &str = "mod soundtrack_settings;";\n',
            "mod soundtrack_setting;\n",
        )
        accepted_owner = (
            "mod soundtrack_settings;\n",
            "mod r#soundtrack_settings {}\n",
            "pub(crate) mod soundtrack_settings;\n",
            "pub(super) mod r#soundtrack_settings {}\n",
            "pub(in crate)\nmod\nsoundtrack_settings\n;\n",
        )
        for declaration, expected in (
            *[
                (
                    declaration,
                    [
                        "focused litchi-keynote soundtrack settings public API is "
                        "missing private package owner module: "
                        "crates/litchi-keynote/src/package.rs"
                    ],
                )
                for declaration in rejected_owner
            ],
            *[(declaration, []) for declaration in accepted_owner],
        ):
            with self.subTest(scope="owner", declaration=declaration):
                with tempfile.TemporaryDirectory() as directory:
                    root = Path(directory)
                    add_keynote_soundtrack_settings_canonical_scaffold(root)
                    package_export = root / boundaries.KEYNOTE_SOUNDTRACK_SETTINGS_EXPORT_SOURCES[1]
                    package_export.write_text(declaration, encoding="utf-8")
                    self.assertEqual(
                        boundaries.audit_keynote_soundtrack_settings_facade_source_topology(
                            root
                        ),
                        expected,
                    )

        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            add_keynote_soundtrack_settings_canonical_scaffold(root)
            (root / boundaries.KEYNOTE_SOUNDTRACK_SETTINGS_OWNER_SOURCE).unlink()
            self.assertEqual(
                boundaries.audit_keynote_soundtrack_settings_facade_source_topology(root),
                [
                    "focused litchi-keynote soundtrack settings public API is "
                    "missing private package owner source: crates/litchi-keynote/"
                    "src/package/soundtrack_settings.rs"
                ],
            )

    def test_focused_keynote_soundtrack_settings_rejects_duplicate_modules(self) -> None:
        for declaration in (
            "pub mod transaction;",
            "pub mod r#transaction;",
            "pub mod transaction {}",
            "pub\nmod\nr#transaction\n{}",
        ):
            with self.subTest(scope="transaction", declaration=declaration):
                with tempfile.TemporaryDirectory() as directory:
                    root = Path(directory)
                    add_keynote_soundtrack_settings_canonical_scaffold(root)
                    semantic = root / boundaries.KEYNOTE_SOUNDTRACK_SETTINGS_SEMANTIC_SOURCE
                    line = semantic.read_text(encoding="utf-8").count("\n") + 1
                    semantic.write_text(
                        semantic.read_text(encoding="utf-8") + declaration + "\n",
                        encoding="utf-8",
                    )
                    self.assertEqual(
                        boundaries.audit_keynote_soundtrack_settings_facade_source_topology(
                            root
                        ),
                        [
                            "focused litchi-keynote soundtrack settings public API "
                            "exposes duplicate soundtrack::transaction module: "
                            f"crates/litchi-keynote/src/soundtrack.rs:{line}"
                        ],
                    )

    def test_focused_keynote_soundtrack_settings_rejects_all_flat_aliases_and_host_setter(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            relative_sources = (
                boundaries.KEYNOTE_SOUNDTRACK_SETTINGS_IMPLEMENTATION_SOURCES
                + boundaries.KEYNOTE_SOUNDTRACK_SETTINGS_EXPORT_SOURCES
            )
            sources = tuple(root / path for path in relative_sources)
            declarations: dict[Path, list[str]] = {path: [] for path in sources}
            expected: list[str] = []
            for index, name in enumerate(
                sorted(boundaries.KEYNOTE_SOUNDTRACK_SETTINGS_FLAT_ALIASES)
            ):
                source_index = index % len(sources)
                path = sources[source_index]
                if source_index == 0:
                    declaration = f"pub struct {name};"
                elif source_index == 1:
                    declaration = f"pub type {name} = bool;"
                else:
                    declaration = f"pub use crate::legacy::Legacy as {name};"
                declarations[path].append(declaration)
                expected.append(
                    "focused litchi-keynote soundtrack settings public API retains "
                    f"flat alias {name}: {relative_sources[source_index]}:"
                    f"{len(declarations[path])}"
                )
            owner = sources[1]
            declarations[owner].append("pub fn set_soundtrack_settings() {}")
            expected.append(
                "focused litchi-keynote soundtrack settings public API retains "
                "host-style public member set_soundtrack_settings: "
                f"{relative_sources[1]}:{len(declarations[owner])}"
            )
            for path, lines in declarations.items():
                path.parent.mkdir(parents=True, exist_ok=True)
                path.write_text("\n".join(lines) + "\n", encoding="utf-8")
            add_keynote_soundtrack_settings_canonical_scaffold(root)

            self.assertEqual(
                boundaries.audit_keynote_soundtrack_settings_facade_source_topology(root),
                sorted(expected),
            )

    def test_focused_keynote_soundtrack_settings_rejects_root_aliases_glob_and_owner_aliases(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            lib_export, package_export = (
                root / path
                for path in boundaries.KEYNOTE_SOUNDTRACK_SETTINGS_EXPORT_SOURCES
            )
            lib_export.parent.mkdir(parents=True)
            aliases = sorted(boundaries.KEYNOTE_SOUNDTRACK_SETTINGS_SHORT_NAMES)
            midpoint = len(aliases) // 2
            lib_aliases = aliases[:midpoint]
            package_aliases = aliases[midpoint:]
            lib_export.write_text(
                "pub use crate::soundtrack::{" + ", ".join(lib_aliases) + "};\n",
                encoding="utf-8",
            )
            package_export.write_text(
                "pub use crate::soundtrack_settings::{"
                + ", ".join(package_aliases)
                + "};\n"
                "pub use crate::soundtrack::*;\n",
                encoding="utf-8",
            )
            add_keynote_soundtrack_settings_canonical_scaffold(root)

            self.assertEqual(
                boundaries.audit_keynote_soundtrack_settings_facade_source_topology(root),
                sorted(
                    [
                        *[
                            "focused litchi-keynote soundtrack settings public API "
                            f"retains root alias {name}: crates/litchi-keynote/src/lib.rs:1"
                            for name in lib_aliases
                        ],
                        *[
                            "focused litchi-keynote soundtrack settings public API "
                            f"retains root alias {name}: crates/litchi-keynote/src/package.rs:1"
                            for name in package_aliases
                        ],
                        "focused litchi-keynote soundtrack settings public API "
                        "exposes public soundtrack owner alias: "
                        "crates/litchi-keynote/src/lib.rs:1",
                        "focused litchi-keynote soundtrack settings public API "
                        "exposes public soundtrack owner alias: "
                        "crates/litchi-keynote/src/package.rs:1",
                        "focused litchi-keynote soundtrack settings public API "
                        "exposes public soundtrack owner alias: "
                        "crates/litchi-keynote/src/package.rs:2",
                        "focused litchi-keynote soundtrack settings public API "
                        "retains root aliases via soundtrack glob: "
                        "crates/litchi-keynote/src/package.rs:2",
                    ]
                ),
            )

    def test_focused_keynote_soundtrack_settings_rejects_public_owner_alias_variants(
        self,
    ) -> None:
        for relative in boundaries.KEYNOTE_SOUNDTRACK_SETTINGS_EXPORT_SOURCES:
            for declaration in (
                "pub use crate::soundtrack as audio;",
                "pub use crate::r#soundtrack_settings as audio;",
                "pub\nuse\ncrate::soundtrack as audio;",
                "pub type soundtrack_settings = bool;",
            ):
                with self.subTest(relative=relative, declaration=declaration):
                    with tempfile.TemporaryDirectory() as directory:
                        root = Path(directory)
                        add_keynote_soundtrack_settings_canonical_scaffold(root)
                        path = root / relative
                        source = path.read_text(encoding="utf-8")
                        line = source.count("\n") + 1
                        path.write_text(source + declaration + "\n", encoding="utf-8")
                        self.assertEqual(
                            boundaries.audit_keynote_soundtrack_settings_facade_source_topology(
                                root
                            ),
                            [
                                "focused litchi-keynote soundtrack settings public "
                                "API exposes public soundtrack owner alias: "
                                f"{relative}:{line}"
                            ],
                        )

    def test_focused_keynote_soundtrack_settings_rejects_physical_and_media_leaks(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            semantic = root / boundaries.KEYNOTE_SOUNDTRACK_SETTINGS_SEMANTIC_SOURCE
            semantic.parent.mkdir(parents=True)
            semantic.write_text(
                "pub fn settings(r#source_bytes: &[u8], r#object_id: u64) "
                "-> DocumentArchive { todo!() }\n"
                "pub type Projection = buffa::DocumentArchiveView;\n"
                "impl prost::Message for soundtrack::Settings {}\n",
                encoding="utf-8",
            )
            owner = root / boundaries.KEYNOTE_SOUNDTRACK_SETTINGS_OWNER_SOURCE
            owner.parent.mkdir(parents=True, exist_ok=True)
            owner_lines = [
                f"pub type Physical{index} = {name};"
                for index, name in enumerate(
                    sorted(boundaries.KEYNOTE_SOUNDTRACK_SETTINGS_PHYSICAL_TYPES)
                )
            ]
            owner_lines.extend(
                f"pub type WireLeak{index} = {name};"
                for index, name in enumerate(
                    sorted(boundaries.KEYNOTE_SOUNDTRACK_SETTINGS_WIRE_TYPES)
                )
            )
            owner_lines.extend(
                (
                    f"pub type MediaLeak{index} = {name};"
                    if name[0].isupper()
                    else f"pub fn {name}() {{}}"
                )
                for index, name in enumerate(
                    sorted(boundaries.KEYNOTE_SOUNDTRACK_SETTINGS_MEDIA_TOPOLOGY_NAMES)
                )
            )
            owner_lines.append(
                "pub type ProtoProjection = (kn::SoundtrackArchive, "
                "tsp::ReferenceArchive);"
            )
            owner.write_text("\n".join(owner_lines) + "\n", encoding="utf-8")
            lib_export, package_export = (
                root / path
                for path in boundaries.KEYNOTE_SOUNDTRACK_SETTINGS_EXPORT_SOURCES
            )
            lib_export.write_text(
                "pub fn edit_soundtrack_settings("
                "value: litchi_iwa_protos::GeneratedSoundtrack) "
                "-> soundtrack::Patch { todo!() }\n",
                encoding="utf-8",
            )
            package_export.write_text(
                "pub fn apply_soundtrack_settings(value: prost_types::MessageInfo) "
                "-> soundtrack::Commit { todo!() }\n"
                "pub fn soundtrack_settings() -> SourceBytes { todo!() }\n",
                encoding="utf-8",
            )
            add_keynote_soundtrack_settings_canonical_scaffold(root)

            violations = boundaries.audit_keynote_soundtrack_settings_facade_source_topology(
                root
            )

            self.assertEqual(violations, sorted(violations))
            self.assertTrue(
                all(
                    violation.startswith(
                        "focused litchi-keynote soundtrack settings public API exposes "
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
                "protobuf type kn",
                "protobuf type ProtoProjection",
                "archive/IWA type SoundtrackArchive",
                "protobuf type tsp",
                "archive/IWA type ReferenceArchive",
                "archive/IWA type litchi_iwa_protos",
                "generated type GeneratedSoundtrack",
                "protobuf type prost_types",
                "archive/IWA type MessageInfo",
                "raw source bytes SourceBytes",
                *tuple(
                    f"archive/IWA type {name}"
                    for name in sorted(
                        boundaries.KEYNOTE_SOUNDTRACK_SETTINGS_PHYSICAL_TYPES
                    )
                ),
                *tuple(
                    f"wire type {name}"
                    for name in sorted(boundaries.KEYNOTE_SOUNDTRACK_SETTINGS_WIRE_TYPES)
                ),
                *tuple(
                    f"soundtrack media topology {name}"
                    for name in sorted(
                        boundaries.KEYNOTE_SOUNDTRACK_SETTINGS_MEDIA_TOPOLOGY_NAMES
                    )
                ),
            )
            self.assertEqual(len(violations), 53)
            for fragment in expected_fragments:
                with self.subTest(fragment=fragment):
                    self.assertTrue(
                        any(fragment in violation for violation in violations),
                        msg=f"missing focused soundtrack-settings leak: {fragment}",
                    )

    def test_focused_keynote_soundtrack_settings_recursively_scans_private_helpers(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            add_keynote_soundtrack_settings_canonical_scaffold(root)
            nested_relative = (
                boundaries.KEYNOTE_SOUNDTRACK_SETTINGS_OWNER_HELPER_ROOT
                / "future"
                / "nested.rs"
            )
            nested = root / nested_relative
            nested.parent.mkdir(parents=True, exist_ok=True)
            nested.write_text(
                "pub fn expose(source_bytes: &[u8], object_id: u64) "
                "-> ArchiveObject { todo!() }\n"
                "pub type SoundtrackEdit = WireView;\n"
                "pub fn media_items() {}\n"
                "impl prost::Message for SoundtrackPatch<RawMessage> {}\n",
                encoding="utf-8",
            )

            violations = boundaries.audit_keynote_soundtrack_settings_facade_source_topology(
                root
            )
            path = str(nested_relative)
            self.assertEqual(
                violations,
                sorted(
                    [
                        "focused litchi-keynote soundtrack settings public API "
                        f"exposes raw source bytes source_bytes: {path}:1",
                        "focused litchi-keynote soundtrack settings public API "
                        f"exposes raw byte slice &[u8]: {path}:1",
                        "focused litchi-keynote soundtrack settings public API "
                        f"exposes raw identifier object_id: {path}:1",
                        "focused litchi-keynote soundtrack settings public API "
                        f"exposes archive/IWA type ArchiveObject: {path}:1",
                        "focused litchi-keynote soundtrack settings public API "
                        f"retains flat alias SoundtrackEdit: {path}:2",
                        "focused litchi-keynote soundtrack settings public API "
                        f"exposes wire type WireView: {path}:2",
                        "focused litchi-keynote soundtrack settings public API "
                        f"exposes soundtrack media topology media_items: {path}:3",
                        "focused litchi-keynote soundtrack settings public API "
                        f"exposes protobuf type prost: {path}:4",
                        "focused litchi-keynote soundtrack settings public API "
                        f"exposes protobuf type Message: {path}:4",
                        "focused litchi-keynote soundtrack settings public API "
                        f"exposes archive/IWA type RawMessage: {path}:4",
                    ]
                ),
            )

    def test_focused_keynote_soundtrack_settings_allows_canonical_and_media_item_surfaces(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            semantic = root / boundaries.KEYNOTE_SOUNDTRACK_SETTINGS_SEMANTIC_SOURCE
            semantic.parent.mkdir(parents=True)
            semantic.write_text(
                "// pub type SoundtrackEdit = DocumentArchive;\n"
                'const NOTE: &str = "pub struct SoundtrackSettingsPatch;";\n'
                "/* pub fn media_items() -> &[u8] {} */\n"
                "pub enum PlaybackMode { Off, PlayOnce, Loop, Unknown(i32) }\n"
                "pub struct PlaybackSettings { pub volume: Option<f64> }\n"
                "pub fn as_raw(mode: PlaybackMode) -> i32 { todo!() }\n"
                "pub fn from_raw(value: i32) -> PlaybackMode { todo!() }\n"
                "pub fn source_fingerprint() -> u64 { 0 }\n"
                "pub fn soundtrack_item_count() -> usize { 0 }\n"
                "fn source_bytes(source_bytes: &[u8], object_id: u64) "
                "-> ArchiveObject { todo!() }\n"
                "pub(crate) fn restricted(value: WireView) {}\n"
                "mod transaction;\n"
                "pub(crate) mod r#transaction;\n",
                encoding="utf-8",
            )
            owner = root / boundaries.KEYNOTE_SOUNDTRACK_SETTINGS_OWNER_SOURCE
            owner.parent.mkdir(parents=True, exist_ok=True)
            private_aliases = [
                ("struct" if index % 2 == 0 else "pub(crate) struct")
                + f" {name};"
                for index, name in enumerate(
                    sorted(boundaries.KEYNOTE_SOUNDTRACK_SETTINGS_FLAT_ALIASES)
                )
            ]
            private_vocabulary = [
                f"pub(super) type Private{index} = {name};"
                for index, name in enumerate(
                    sorted(
                        boundaries.KEYNOTE_SOUNDTRACK_SETTINGS_PHYSICAL_TYPES
                        | boundaries.KEYNOTE_SOUNDTRACK_SETTINGS_WIRE_TYPES
                        | boundaries.KEYNOTE_SOUNDTRACK_SETTINGS_MEDIA_TOPOLOGY_NAMES
                    )
                )
            ]
            owner.write_text(
                "\n".join(
                    [
                        "pub fn soundtrack_settings() -> soundtrack::Settings { todo!() }",
                        "pub fn edit_soundtrack_settings() "
                        "-> soundtrack::Edit { todo!() }",
                        "pub fn apply_soundtrack_settings(patch: soundtrack::Patch) "
                        "-> soundtrack::Commit { todo!() }",
                        *private_aliases,
                        *private_vocabulary,
                        "fn set_soundtrack_settings() {}",
                        "impl SoundtrackEdit {}",
                    ]
                )
                + "\n",
                encoding="utf-8",
            )
            lib_export, package_export = (
                root / path
                for path in boundaries.KEYNOTE_SOUNDTRACK_SETTINGS_EXPORT_SOURCES
            )
            private_roots = "\n".join(
                ("use" if index % 2 == 0 else "pub(crate) use")
                + f" crate::soundtrack::{name};"
                for index, name in enumerate(
                    sorted(boundaries.KEYNOTE_SOUNDTRACK_SETTINGS_SHORT_NAMES)
                )
            )
            lib_export.write_text(
                "// pub use crate::soundtrack::{Mode, Settings};\n"
                'const GLOB_NOTE: &str = "pub use crate::soundtrack::*;";\n'
                + private_roots
                + "\npub(crate) use crate::soundtrack::*;\n"
                "pub use crate::animation::{Mode, Settings, Edit, Patch, Commit, "
                "Diagnostics, Error, LimitKind};\n"
                "pub struct KeynoteSoundtrackItemInfo;\n",
                encoding="utf-8",
            )
            package_export.write_text(
                "mod soundtrack_settings;\n"
                "pub(crate) mod r#soundtrack_settings;\n"
                "pub(super) use crate::soundtrack::{Mode, Settings};\n"
                "pub fn soundtrack_items() -> Vec<KeynoteSoundtrackItemInfo> { todo!() }\n",
                encoding="utf-8",
            )
            helper = (
                root
                / boundaries.KEYNOTE_SOUNDTRACK_SETTINGS_OWNER_HELPER_ROOT
                / "private.rs"
            )
            helper.parent.mkdir(parents=True, exist_ok=True)
            helper.write_text(
                "fn source_bytes(source_bytes: &[u8], object_id: u64) "
                "-> ArchiveObject { todo!() }\n"
                "pub(crate) type SoundtrackPatch = WireView;\n"
                "pub(super) fn media_items() {}\n"
                "impl SoundtrackEdit {}\n",
                encoding="utf-8",
            )
            non_rust = helper.with_suffix(".txt")
            non_rust.write_text(
                "pub type SoundtrackEdit = ArchiveObject;\n"
                "pub fn media_items(value: &[u8]) {}\n",
                encoding="utf-8",
            )
            nonfocused = root / boundaries.KEYNOTE_SOURCE_ROOT / "soundtrack_items.rs"
            nonfocused.write_text(
                "\n".join(
                    f"pub struct {name};"
                    for name in sorted(boundaries.KEYNOTE_SOUNDTRACK_SETTINGS_FLAT_ALIASES)
                )
                + "\npub struct KeynoteSoundtrackItemInfo;\n"
                "pub fn media_items() -> Vec<DataReference> { todo!() }\n"
                "pub fn set_soundtrack_settings() {}\n"
                "pub fn item(object_id: u64) -> ArchiveObject { todo!() }\n",
                encoding="utf-8",
            )
            other_owner = root / "crates/litchi-pages/src/soundtrack.rs"
            other_owner.parent.mkdir(parents=True)
            other_owner.write_text(
                "pub struct SoundtrackSettings;\n"
                "pub fn set_soundtrack_settings() {}\n"
                "pub fn media_items() -> Vec<DataReference> { todo!() }\n",
                encoding="utf-8",
            )
            add_keynote_soundtrack_settings_canonical_scaffold(root)

            self.assertEqual(
                boundaries.audit_keynote_soundtrack_settings_facade_source_topology(root),
                [],
            )

        for declaration in (
            "pub mod soundtrack_settings;",
            "pub mod r#soundtrack_settings;",
            "pub mod soundtrack_settings {}",
            "pub\nmod\nr#soundtrack_settings\n{}",
        ):
            with self.subTest(scope="package", declaration=declaration):
                with tempfile.TemporaryDirectory() as directory:
                    root = Path(directory)
                    add_keynote_soundtrack_settings_canonical_scaffold(root)
                    package_export = root / boundaries.KEYNOTE_SOUNDTRACK_SETTINGS_EXPORT_SOURCES[1]
                    package_export.write_text(declaration + "\n", encoding="utf-8")
                    self.assertEqual(
                        boundaries.audit_keynote_soundtrack_settings_facade_source_topology(
                            root
                        ),
                        [
                            "focused litchi-keynote soundtrack settings public API "
                            "exposes duplicate package::soundtrack_settings module: "
                            "crates/litchi-keynote/src/package.rs:1"
                        ],
                    )

    def test_keynote_slide_transition_boundary_inventories_are_exact(self) -> None:
        self.assertEqual(
            boundaries.RETIRED_IWA_KEYNOTE_SLIDE_TRANSITION_METHODS,
            ("slide_transition", "set_slide_transition", "clear_slide_transition"),
        )
        self.assertEqual(
            boundaries.RETIRED_IWA_KEYNOTE_SLIDE_TRANSITION_SOURCES,
            (Path("crates/litchi-iwa/src/keynote/editor/transition_lifecycle.rs"),),
        )
        self.assertEqual(
            boundaries.RETIRED_IWA_KEYNOTE_SLIDE_TRANSITION_MODULES,
            ("transition_lifecycle",),
        )
        self.assertEqual(
            boundaries.RETIRED_IWA_KEYNOTE_SLIDE_TRANSITION_EXAMPLES,
            (
                Path("crates/litchi-iwa/examples/clear_keynote_transition.rs"),
                Path("crates/litchi-iwa/examples/edit_keynote_transition.rs"),
                Path("crates/litchi-iwa/examples/set_keynote_transition_effect.rs"),
            ),
        )
        self.assertEqual(
            boundaries.KEYNOTE_SLIDE_TRANSITION_IMPLEMENTATION_SOURCES,
            (
                Path("crates/litchi-keynote/src/transition.rs"),
                Path("crates/litchi-keynote/src/package/slide_transition.rs"),
            ),
        )
        self.assertEqual(
            boundaries.KEYNOTE_SLIDE_TRANSITION_EXPORT_SOURCES,
            (
                Path("crates/litchi-keynote/src/lib.rs"),
                Path("crates/litchi-keynote/src/package.rs"),
            ),
        )
        self.assertEqual(
            boundaries.KEYNOTE_SLIDE_TRANSITION_CANONICAL_TYPES,
            ("Edit", "Patch", "Commit", "Diagnostics", "Error", "LimitKind"),
        )
        self.assertEqual(
            boundaries.KEYNOTE_SLIDE_TRANSITION_SEMANTIC_TYPES,
            (
                "Acceleration",
                "AccelerationKind",
                "AnimationParameters",
                "CustomParameters",
                "Direction",
                "Effect",
                "MosaicType",
                "Settings",
                "SettingsBuilder",
                "TextDelivery",
                "TextDeliveryKind",
                "TimingCurveSlot",
            ),
        )
        self.assertEqual(
            boundaries.KEYNOTE_SLIDE_TRANSITION_FLAT_ALIAS_PREFIXES,
            ("SlideTransition", "Transition"),
        )
        self.assertEqual(
            boundaries.KEYNOTE_SLIDE_TRANSITION_FLAT_ALIASES,
            frozenset(
                prefix + suffix
                for prefix in ("SlideTransition", "Transition")
                for suffix in (
                    "Edit",
                    "Patch",
                    "Commit",
                    "Diagnostics",
                    "Error",
                    "LimitKind",
                )
            ),
        )
        self.assertEqual(
            boundaries.KEYNOTE_SLIDE_TRANSITION_ROOT_ALIASES,
            frozenset(
                boundaries.KEYNOTE_SLIDE_TRANSITION_SEMANTIC_TYPES
                + boundaries.KEYNOTE_SLIDE_TRANSITION_CANONICAL_TYPES
            ),
        )
        self.assertEqual(
            boundaries.KEYNOTE_SLIDE_TRANSITION_PHYSICAL_TYPES,
            frozenset(
                {
                    "Archive",
                    "EntryEdit",
                    "IWorkPackage",
                    "PhysicalSource",
                    "RawMessage",
                    "SnappyStream",
                    "SourceCatalog",
                    "TransitionSettingsSnapshot",
                }
            ),
        )
        self.assertEqual(
            boundaries.KEYNOTE_SLIDE_TRANSITION_WIRE_TYPES,
            frozenset(
                {
                    "WireDescent",
                    "WireError",
                    "WireLimits",
                    "WireResourceLimit",
                    "WireView",
                }
            ),
        )
        self.assertEqual(
            boundaries.KEYNOTE_SLIDE_TRANSITION_SEMANTIC_IDENTIFIER_NAMES,
            frozenset(
                {"MAX_IDENTIFIER_BYTES", "from_identifier", "identifier", "identifiers"}
            ),
        )
        self.assertEqual(
            boundaries.KEYNOTE_SLIDE_TRANSITION_SEMANTIC_OPAQUE_PAYLOAD_MEMBERS,
            frozenset(
                {
                    "color_payload",
                    "set_color_payload",
                    "set_timing_curve_payload",
                    "timing_curve_payload",
                    "timing_curve_payloads",
                }
            ),
        )

    def test_retired_iwa_keynote_slide_transition_surface_cannot_return(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            retained = root / boundaries.IWA_KEYNOTE_SOURCE_ROOT / "editor/transition.rs"
            retained.parent.mkdir(parents=True)
            retained.write_text(
                "pub(in crate::keynote) async unsafe fn set_slide_transition() {}\n",
                encoding="utf-8",
            )
            nested = root / boundaries.IWA_KEYNOTE_SOURCE_ROOT / "legacy/transition.rs"
            nested.parent.mkdir(parents=True)
            nested.write_text(
                "fn r#slide_transition() {}\n"
                "pub(super) const fn clear_slide_transition() {}\n",
                encoding="utf-8",
            )
            editor = root / boundaries.IWA_KEYNOTE_EDITOR_SOURCE
            editor.write_text(
                "pub(crate) mod r#transition_lifecycle;\n",
                encoding="utf-8",
            )
            for retired in boundaries.RETIRED_IWA_KEYNOTE_SLIDE_TRANSITION_SOURCES:
                path = root / retired
                path.parent.mkdir(parents=True, exist_ok=True)
                path.write_text("// retired owner returned\n", encoding="utf-8")
            for retired in boundaries.RETIRED_IWA_KEYNOTE_SLIDE_TRANSITION_EXAMPLES:
                path = root / retired
                path.parent.mkdir(parents=True, exist_ok=True)
                path.write_text("// retired example returned\n", encoding="utf-8")

            self.assertEqual(
                boundaries.audit_iwa_keynote_slide_transition_source_topology(root),
                sorted(
                    [
                        "retired litchi-iwa Keynote slide-transition method "
                        "set_slide_transition: "
                        "crates/litchi-iwa/src/keynote/editor/transition.rs:1",
                        "retired litchi-iwa Keynote slide-transition method "
                        "slide_transition: "
                        "crates/litchi-iwa/src/keynote/legacy/transition.rs:1",
                        "retired litchi-iwa Keynote slide-transition method "
                        "clear_slide_transition: "
                        "crates/litchi-iwa/src/keynote/legacy/transition.rs:2",
                        "retired litchi-iwa Keynote slide-transition module "
                        "transition_lifecycle: "
                        "crates/litchi-iwa/src/keynote/editor.rs:1",
                        *[
                            "retired litchi-iwa Keynote slide-transition source "
                            f"returned: {path}"
                            for path in boundaries.RETIRED_IWA_KEYNOTE_SLIDE_TRANSITION_SOURCES
                        ],
                        *[
                            "retired litchi-iwa Keynote slide-transition example "
                            f"returned: {path}"
                            for path in boundaries.RETIRED_IWA_KEYNOTE_SLIDE_TRANSITION_EXAMPLES
                        ],
                    ]
                ),
            )

    def test_retired_iwa_keynote_slide_transition_module_variants(self) -> None:
        for declaration in (
            "mod transition_lifecycle;",
            "pub(crate) mod transition_lifecycle;",
            "pub(super) mod r#transition_lifecycle {}",
            "pub(in crate::keynote)\nmod\nr#transition_lifecycle\n{}",
            "pub mod transition_lifecycle { pub struct Legacy; }",
        ):
            with self.subTest(declaration=declaration):
                with tempfile.TemporaryDirectory() as directory:
                    root = Path(directory)
                    editor = root / boundaries.IWA_KEYNOTE_EDITOR_SOURCE
                    editor.parent.mkdir(parents=True)
                    editor.write_text(declaration + "\n", encoding="utf-8")

                    self.assertEqual(
                        boundaries.audit_iwa_keynote_slide_transition_source_topology(
                            root
                        ),
                        [
                            "retired litchi-iwa Keynote slide-transition module "
                            "transition_lifecycle: "
                            "crates/litchi-iwa/src/keynote/editor.rs:1"
                        ],
                    )

    def test_retired_iwa_keynote_slide_transition_readme_calls_and_examples(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            readme = root / boundaries.IWA_KEYNOTE_README
            readme.parent.mkdir(parents=True)
            readme.write_text(
                "\n".join(
                    [
                        "let value = keynote",
                        "    .",
                        "    r#slide_transition",
                        "    (",
                        "editor.r#set_slide_transition(",
                        "reopened.r#clear_slide_transition(",
                        "crate::keynote::KeynoteEditor",
                        "    ::",
                        "    slide_transition(",
                        "r#KeynoteEditor::r#set_slide_transition(",
                        "other.clear_slide_transition(",
                        "Run `clear_keynote_transition`.",
                        "Run edit_keynote_transition.rs.",
                        "Run set_keynote_transition_effect.",
                    ]
                )
                + "\n",
                encoding="utf-8",
            )

            self.assertEqual(
                boundaries.audit_iwa_keynote_slide_transition_source_topology(root),
                sorted(
                    [
                        "retired litchi-iwa Keynote slide-transition README call "
                        "slide_transition: crates/litchi-iwa/README.md:3",
                        "retired litchi-iwa Keynote slide-transition README call "
                        "set_slide_transition: crates/litchi-iwa/README.md:5",
                        "retired litchi-iwa Keynote slide-transition README call "
                        "clear_slide_transition: crates/litchi-iwa/README.md:6",
                        "retired litchi-iwa Keynote slide-transition README call "
                        "slide_transition: crates/litchi-iwa/README.md:9",
                        "retired litchi-iwa Keynote slide-transition README call "
                        "set_slide_transition: crates/litchi-iwa/README.md:10",
                        "retired litchi-iwa Keynote slide-transition README call "
                        "clear_slide_transition: crates/litchi-iwa/README.md:11",
                        "retired litchi-iwa Keynote slide-transition README example "
                        "reference clear_keynote_transition: "
                        "crates/litchi-iwa/README.md:12",
                        "retired litchi-iwa Keynote slide-transition README example "
                        "reference edit_keynote_transition: "
                        "crates/litchi-iwa/README.md:13",
                        "retired litchi-iwa Keynote slide-transition README example "
                        "reference set_keynote_transition_effect: "
                        "crates/litchi-iwa/README.md:14",
                    ]
                ),
            )

    def test_iwa_keynote_slide_transition_policy_ignores_safe_surfaces(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            host = root / boundaries.IWA_KEYNOTE_SOURCE_ROOT / "legacy/transition_old.rs"
            host.parent.mkdir(parents=True)
            host.write_text(
                "\n".join(
                    [
                        "// pub fn slide_transition() {}",
                        'const NOTE: &str = "fn set_slide_transition() {}";',
                        "/* fn clear_slide_transition() {}",
                        "   /* fn slide_transition() {} */",
                        "   fn set_slide_transition() {} */",
                        'const RAW_NOTE: &str = r###"fn clear_slide_transition() {}"###;',
                        "pub fn slide_transition_snapshot() {}",
                        "pub fn set_slide_transitions() {}",
                        "pub fn clear_slide_transition_cache() {}",
                    ]
                )
                + "\n",
                encoding="utf-8",
            )
            editor = root / boundaries.IWA_KEYNOTE_EDITOR_SOURCE
            editor.write_text(
                "\n".join(
                    [
                        "// mod transition_lifecycle;",
                        'const NOTE: &str = "mod transition_lifecycle;";',
                        "/* pub mod r#transition_lifecycle {} */",
                        "mod transition_lifecycle_old;",
                        "pub use litchi_keynote::transition::{Effect, Acceleration};",
                        "pub use crate::transition::Effect;",
                    ]
                )
                + "\n",
                encoding="utf-8",
            )
            readme = root / boundaries.IWA_KEYNOTE_README
            readme.write_text(
                "\n".join(
                    [
                        "The slide_transition method is retired.",
                        "Use `slide_transition` rather than `.slide_transition`.",
                        "slide_transition(",
                        "set_slide_transition(",
                        "keynote.slide_transition_snapshot(",
                        "editor.set_slide_transitions(",
                        "package.slide_transition(",
                        "package.edit_slide_transition(",
                        "edit.clear(",
                        "create_keynote_transition.rs",
                        "clear_keynote_transitions.rs",
                    ]
                )
                + "\n",
                encoding="utf-8",
            )
            declarations = "\n".join(
                f"pub fn {name}() {{}}"
                for name in boundaries.RETIRED_IWA_KEYNOTE_SLIDE_TRANSITION_METHODS
            ) + "\n"
            for relative in (
                Path("crates/litchi-keynote/src/package/slide_transition.rs"),
                Path("crates/litchi-iwa/src/pages/editor/transition.rs"),
            ):
                path = root / relative
                path.parent.mkdir(parents=True, exist_ok=True)
                path.write_text(declarations, encoding="utf-8")
            non_rust = root / boundaries.IWA_KEYNOTE_SOURCE_ROOT / "transition.txt"
            non_rust.write_text(declarations, encoding="utf-8")
            retained_example = root / "crates/litchi-iwa/examples/create_keynote_transition.rs"
            retained_example.parent.mkdir(parents=True, exist_ok=True)
            retained_example.write_text("// retained creation example\n", encoding="utf-8")
            for relative in (Path("README.md"), Path("crates/litchi-keynote/README.md")):
                other = root / relative
                other.parent.mkdir(parents=True, exist_ok=True)
                other.write_text(
                    "keynote.slide_transition(\nclear_keynote_transition\n",
                    encoding="utf-8",
                )

            self.assertEqual(
                boundaries.audit_iwa_keynote_slide_transition_source_topology(root),
                [],
            )

    def test_focused_keynote_slide_transition_public_api_rejects_physical_leaks(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            semantic, transaction = (
                root / path
                for path in boundaries.KEYNOTE_SLIDE_TRANSITION_IMPLEMENTATION_SOURCES
            )
            semantic.parent.mkdir(parents=True)
            semantic.write_text(
                "pub fn transition(r#source_bytes: &[u8], r#object_id: u64) "
                "-> (DocumentArchive, IWorkPackage, SourceCatalog) {}\n"
                "pub type Effect = buffa::DocumentArchiveView;\n"
                "impl prost::Message for transition::Settings {}\n",
                encoding="utf-8",
            )
            transaction.parent.mkdir(parents=True, exist_ok=True)
            transaction.write_text(
                "\n".join(
                    [
                        "pub type Edit = Archive;",
                        "pub type Patch = EntryEdit;",
                        "pub type Commit = IWorkPackage;",
                        "pub type Diagnostics = PhysicalSource;",
                        "pub type Error = RawMessage;",
                        "pub type LimitKind = SnappyStream;",
                        "pub type TransitionPhysical = SourceCatalog;",
                        "pub type TransitionSnapshot = TransitionSettingsSnapshot;",
                        "pub type TransitionDescent = wire::WireDescent;",
                        "pub type TransitionWireError = wire::WireError;",
                        "pub type TransitionWireLimits = WireLimits;",
                        "pub type TransitionWireResource = WireResourceLimit;",
                        "pub type TransitionWireView = WireView;",
                        "pub type TransitionNative = NativeObject;",
                    ]
                )
                + "\n",
                encoding="utf-8",
            )
            lib_export, package_export = (
                root / path
                for path in boundaries.KEYNOTE_SLIDE_TRANSITION_EXPORT_SOURCES
            )
            lib_export.write_text(
                "pub fn edit_slide_transition("
                "value: litchi_iwa_protos::GeneratedTransition) "
                "-> transition::Patch {}\n",
                encoding="utf-8",
            )
            package_export.write_text(
                "pub fn apply_slide_transition(value: prost_types::MessageInfo) "
                "-> transition::Commit {}\n",
                encoding="utf-8",
            )
            add_keynote_slide_transition_canonical_scaffold(root)

            violations = (
                boundaries.audit_keynote_slide_transition_facade_source_topology(root)
            )

            self.assertEqual(violations, sorted(violations))
            self.assertTrue(
                all(
                    violation.startswith(
                        "focused litchi-keynote slide-transition public API exposes "
                    )
                    for violation in violations
                )
            )
            expected_fragments = (
                "raw source bytes source_bytes",
                "raw byte slice &[u8]",
                "raw identifier object_id",
                "archive/IWA type DocumentArchive",
                "archive/IWA type IWorkPackage",
                "archive/IWA type SourceCatalog",
                "protobuf type buffa",
                "archive/IWA type DocumentArchiveView",
                "protobuf type prost",
                "protobuf type Message",
                "archive/IWA type Archive",
                "archive/IWA type EntryEdit",
                "archive/IWA type PhysicalSource",
                "archive/IWA type RawMessage",
                "archive/IWA type SnappyStream",
                "archive/IWA type TransitionSettingsSnapshot",
                "wire type wire",
                "wire type WireDescent",
                "wire type WireError",
                "wire type WireLimits",
                "wire type WireResourceLimit",
                "wire type WireView",
                "native object NativeObject",
                "archive/IWA type litchi_iwa_protos",
                "generated type GeneratedTransition",
                "protobuf type prost_types",
                "archive/IWA type MessageInfo",
            )
            self.assertEqual(len(violations), 30)
            for fragment in expected_fragments:
                with self.subTest(fragment=fragment):
                    self.assertTrue(
                        any(fragment in violation for violation in violations),
                        msg=f"missing focused slide-transition leak: {fragment}",
                    )

    def test_focused_keynote_slide_transition_requires_each_canonical_type(
        self,
    ) -> None:
        for missing in boundaries.KEYNOTE_SLIDE_TRANSITION_CANONICAL_TYPES:
            with self.subTest(missing=missing):
                with tempfile.TemporaryDirectory() as directory:
                    root = Path(directory)
                    semantic = (
                        root
                        / boundaries.KEYNOTE_SLIDE_TRANSITION_IMPLEMENTATION_SOURCES[0]
                    )
                    semantic.parent.mkdir(parents=True)
                    semantic.write_text(
                        "".join(
                            f"pub struct {name};\n"
                            for name in boundaries.KEYNOTE_SLIDE_TRANSITION_CANONICAL_TYPES
                            if name != missing
                        ),
                        encoding="utf-8",
                    )
                    lib_export = (
                        root / boundaries.KEYNOTE_SLIDE_TRANSITION_EXPORT_SOURCES[0]
                    )
                    lib_export.write_text("pub mod transition;\n", encoding="utf-8")

                    self.assertEqual(
                        boundaries.audit_keynote_slide_transition_facade_source_topology(
                            root
                        ),
                        [
                            "focused litchi-keynote slide-transition public API is "
                            f"missing canonical transition type {missing}: "
                            "crates/litchi-keynote/src/transition.rs"
                        ],
                    )

    def test_focused_keynote_slide_transition_root_module_visibility(self) -> None:
        rejected = (
            "",
            "mod transition;\n",
            "pub(crate) mod transition;\n",
            "pub(super) mod r#transition {}\n",
            "pub(in crate) mod transition;\n",
            "// pub mod transition;\n",
            'const NOTE: &str = "pub mod transition;";\n',
            'const RAW_NOTE: &str = r#"pub mod r#transition {}"#;\n',
        )
        accepted = (
            "pub mod transition;\n",
            "pub mod r#transition;\n",
            "pub mod transition {}\n",
            "pub\nmod\nr#transition\n{}\n",
        )
        for declaration, expected in (
            *[
                (
                    declaration,
                    [
                        "focused litchi-keynote slide-transition public API is "
                        "missing canonical root transition module: "
                        "crates/litchi-keynote/src/lib.rs"
                    ],
                )
                for declaration in rejected
            ],
            *[(declaration, []) for declaration in accepted],
        ):
            with self.subTest(declaration=declaration, expected=expected):
                with tempfile.TemporaryDirectory() as directory:
                    root = Path(directory)
                    semantic = (
                        root
                        / boundaries.KEYNOTE_SLIDE_TRANSITION_IMPLEMENTATION_SOURCES[0]
                    )
                    semantic.parent.mkdir(parents=True)
                    semantic.write_text(
                        "".join(
                            f"pub struct {name};\n"
                            for name in boundaries.KEYNOTE_SLIDE_TRANSITION_CANONICAL_TYPES
                        ),
                        encoding="utf-8",
                    )
                    lib_export = (
                        root / boundaries.KEYNOTE_SLIDE_TRANSITION_EXPORT_SOURCES[0]
                    )
                    lib_export.write_text(declaration, encoding="utf-8")

                    self.assertEqual(
                        boundaries.audit_keynote_slide_transition_facade_source_topology(
                            root
                        ),
                        expected,
                    )

    def test_focused_keynote_slide_transition_api_allows_nested_semantic_surface(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            semantic, transaction = (
                root / path
                for path in boundaries.KEYNOTE_SLIDE_TRANSITION_IMPLEMENTATION_SOURCES
            )
            semantic.parent.mkdir(parents=True)
            semantic.write_text(
                "\n".join(
                    [
                        "// pub type SlideTransitionEdit = DocumentArchive;",
                        'const NOTE: &str = "pub struct TransitionPatch;";',
                        "/* pub type SlideTransitionCommit = buffa::ArchiveView; */",
                        *[
                            f"pub struct {name};"
                            for name in boundaries.KEYNOTE_SLIDE_TRANSITION_SEMANTIC_TYPES
                        ],
                        "pub const MAX_IDENTIFIER_BYTES: usize = 64;",
                        "pub fn from_identifier(identifier: &str, identifiers: &[String]) "
                        "-> Effect { todo!() }",
                        "pub fn color_payload(&self) -> &[u8] { todo!() }",
                        "pub fn set_color_payload(&mut self, payload: &[u8]) {}",
                        "pub fn timing_curve_payload(&self) -> &[u8] { todo!() }",
                        "pub fn timing_curve_payloads(&self) -> &[u8] { todo!() }",
                        "pub fn set_timing_curve_payload(&mut self, payload: &[u8]) {}",
                        "pub struct TransitionSnapshot;",
                        "pub struct SlideTransitions;",
                        "fn source_bytes(source_bytes: &[u8], object_id: u64) "
                        "-> SourceBytes { todo!() }",
                        "pub(crate) fn restricted(archive: DocumentArchive) {}",
                        "impl SlideTransitionEdit {}",
                        "impl prost::Message for Unrelated {}",
                    ]
                )
                + "\n",
                encoding="utf-8",
            )
            private_aliases = [
                (
                    "struct" if index % 2 == 0 else "pub(crate) struct"
                )
                + f" {name};"
                for index, name in enumerate(
                    sorted(boundaries.KEYNOTE_SLIDE_TRANSITION_FLAT_ALIASES)
                )
            ]
            transaction.parent.mkdir(parents=True, exist_ok=True)
            transaction.write_text(
                "\n".join(
                    [
                        *[
                            f"pub struct {name};"
                            for name in boundaries.KEYNOTE_SLIDE_TRANSITION_CANONICAL_TYPES
                        ],
                        *private_aliases,
                        "impl TransitionPatch {}",
                    ]
                )
                + "\n",
                encoding="utf-8",
            )
            lib_export, package_export = (
                root / path
                for path in boundaries.KEYNOTE_SLIDE_TRANSITION_EXPORT_SOURCES
            )
            private_roots = "\n".join(
                ("use" if index % 2 == 0 else "pub(crate) use")
                + f" crate::transition::{name};"
                for index, name in enumerate(
                    sorted(boundaries.KEYNOTE_SLIDE_TRANSITION_ROOT_ALIASES)
                )
            )
            safe_exports = (
                "pub mod transition;\n"
                "// pub use crate::transition::{Effect, Edit};\n"
                'const NOTE: &str = "pub use crate::transition::*;";\n'
                "pub fn edit_slide_transition(settings: transition::Settings, "
                "edit: transition::Edit) -> Result<transition::Commit, "
                "transition::Error> { todo!() }\n"
                + private_roots
                + "\npub(crate) use crate::transition::*;\n"
                "pub use crate::animation::{Acceleration, AccelerationKind, "
                "AnimationParameters, CustomParameters, Direction, Effect, MosaicType, "
                "Settings, SettingsBuilder, TextDelivery, TextDeliveryKind, "
                "TimingCurveSlot, Edit, Patch, Commit, Diagnostics, Error, LimitKind};\n"
                "pub use crate::animation::*;\n"
                "pub fn unrelated(object_id: u64) -> DocumentArchive { todo!() }\n"
            )
            lib_export.write_text(safe_exports, encoding="utf-8")
            package_export.write_text(
                safe_exports.replace("pub mod transition;\n", "")
                + "// pub mod slide_transition;\n"
                + 'const MODULE_NOTE: &str = "pub mod slide_transition;";\n'
                + 'const RAW_NOTE: &str = r#"pub mod r#slide_transition {}"#;\n'
                + "/* pub mod r#slide_transition {} */\n"
                + "mod slide_transition;\n"
                + "pub(crate) mod r#slide_transition;\n"
                + "pub(super) mod slide_transition {}\n"
                + "pub(in crate) mod slide_transition;\n",
                encoding="utf-8",
            )
            nonfocused = root / boundaries.KEYNOTE_SOURCE_ROOT / "slide.rs"
            nonfocused.write_text(
                "\n".join(
                    f"pub struct {name};"
                    for name in sorted(
                        boundaries.KEYNOTE_SLIDE_TRANSITION_FLAT_ALIASES
                    )
                )
                + "\npub fn transition(object_id: u64) -> DocumentArchive { todo!() }\n",
                encoding="utf-8",
            )
            other_owner = root / "crates/litchi-pages/src/transition.rs"
            other_owner.parent.mkdir(parents=True)
            other_owner.write_text(
                "\n".join(
                    f"pub struct {name};"
                    for name in sorted(
                        boundaries.KEYNOTE_SLIDE_TRANSITION_FLAT_ALIASES
                    )
                )
                + "\npub use crate::transition::{Effect, Settings, Edit, Patch};\n"
                "pub fn transition(object_id: u64) -> DocumentArchive { todo!() }\n",
                encoding="utf-8",
            )
            add_keynote_slide_transition_canonical_scaffold(root)

            self.assertEqual(
                boundaries.audit_keynote_slide_transition_facade_source_topology(root),
                [],
            )

    def test_focused_keynote_slide_transition_rejects_all_flat_aliases(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            relative_sources = (
                boundaries.KEYNOTE_SLIDE_TRANSITION_IMPLEMENTATION_SOURCES
                + boundaries.KEYNOTE_SLIDE_TRANSITION_EXPORT_SOURCES
            )
            sources = tuple(root / path for path in relative_sources)
            declarations: dict[Path, list[str]] = {path: [] for path in sources}
            expected: list[str] = []
            for index, name in enumerate(
                sorted(boundaries.KEYNOTE_SLIDE_TRANSITION_FLAT_ALIASES)
            ):
                source_index = index % len(sources)
                path = sources[source_index]
                if source_index == 0:
                    declaration = f"pub struct {name};"
                elif source_index == 1:
                    declaration = f"pub type {name} = Edit;"
                else:
                    declaration = f"pub use crate::legacy::Legacy as {name};"
                declarations[path].append(declaration)
                expected.append(
                    "focused litchi-keynote slide-transition public API retains "
                    f"flat alias {name}: {relative_sources[source_index]}:"
                    f"{len(declarations[path])}"
                )
            for path, lines in declarations.items():
                path.parent.mkdir(parents=True, exist_ok=True)
                path.write_text("\n".join(lines) + "\n", encoding="utf-8")
            add_keynote_slide_transition_canonical_scaffold(root)

            self.assertEqual(
                boundaries.audit_keynote_slide_transition_facade_source_topology(root),
                sorted(expected),
            )

    def test_focused_keynote_slide_transition_exports_reject_all_root_aliases(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            lib_export, package_export = (
                root / path
                for path in boundaries.KEYNOTE_SLIDE_TRANSITION_EXPORT_SOURCES
            )
            lib_export.parent.mkdir(parents=True)
            aliases = sorted(boundaries.KEYNOTE_SLIDE_TRANSITION_ROOT_ALIASES)
            midpoint = len(aliases) // 2
            lib_aliases = aliases[:midpoint]
            package_aliases = aliases[midpoint:]
            lib_export.write_text(
                "pub use crate::transition::{" + ", ".join(lib_aliases) + "};\n",
                encoding="utf-8",
            )
            package_export.write_text(
                "pub use crate::slide_transition::{"
                + ", ".join(package_aliases)
                + "};\n"
                "pub use crate::transition::*;\n",
                encoding="utf-8",
            )
            add_keynote_slide_transition_canonical_scaffold(root)

            self.assertEqual(
                boundaries.audit_keynote_slide_transition_facade_source_topology(root),
                sorted(
                    [
                        *[
                            "focused litchi-keynote slide-transition public API "
                            f"retains root alias {name}: "
                            "crates/litchi-keynote/src/lib.rs:1"
                            for name in lib_aliases
                        ],
                        *[
                            "focused litchi-keynote slide-transition public API "
                            f"retains root alias {name}: "
                            "crates/litchi-keynote/src/package.rs:1"
                            for name in package_aliases
                        ],
                        "focused litchi-keynote slide-transition public API retains "
                        "root aliases via transition glob: "
                        "crates/litchi-keynote/src/package.rs:2",
                    ]
                ),
            )

    def test_focused_keynote_slide_transition_rejects_duplicate_package_module(
        self,
    ) -> None:
        for declaration in (
            "pub mod slide_transition;",
            "pub mod r#slide_transition;",
            "pub mod slide_transition {}",
            "pub\nmod\nr#slide_transition\n{}",
        ):
            with self.subTest(declaration=declaration):
                with tempfile.TemporaryDirectory() as directory:
                    root = Path(directory)
                    package_export = (
                        root / boundaries.KEYNOTE_SLIDE_TRANSITION_EXPORT_SOURCES[1]
                    )
                    package_export.parent.mkdir(parents=True)
                    package_export.write_text(declaration + "\n", encoding="utf-8")
                    add_keynote_slide_transition_canonical_scaffold(root)

                    self.assertEqual(
                        boundaries.audit_keynote_slide_transition_facade_source_topology(
                            root
                        ),
                        [
                            "focused litchi-keynote slide-transition public API exposes "
                            "duplicate package::slide_transition module: "
                            "crates/litchi-keynote/src/package.rs:1"
                        ],
                    )

    def test_keynote_slide_delete_boundary_inventories_are_exact(self) -> None:
        self.assertEqual(
            boundaries.RETIRED_IWA_KEYNOTE_SLIDE_DELETE_METHODS,
            ("remove_slide",),
        )
        self.assertEqual(
            boundaries.RETIRED_IWA_KEYNOTE_SLIDE_DELETE_SOURCE,
            Path("crates/litchi-iwa/src/keynote/editor/slide_delete.rs"),
        )
        self.assertEqual(
            boundaries.RETIRED_IWA_KEYNOTE_SLIDE_DELETE_MODULES,
            ("slide_delete",),
        )
        self.assertEqual(
            boundaries.RETIRED_IWA_KEYNOTE_SLIDE_DELETE_EXAMPLE,
            Path("crates/litchi-iwa/examples/remove_keynote_slide.rs"),
        )
        self.assertEqual(
            boundaries.KEYNOTE_SLIDE_DELETE_IMPLEMENTATION_SOURCES,
            (
                Path("crates/litchi-keynote/src/slide/delete.rs"),
                Path("crates/litchi-keynote/src/package/slide_delete.rs"),
            ),
        )
        self.assertEqual(
            boundaries.KEYNOTE_SLIDE_DELETE_EXPORT_SOURCES,
            (
                Path("crates/litchi-keynote/src/lib.rs"),
                Path("crates/litchi-keynote/src/package.rs"),
                Path("crates/litchi-keynote/src/slide.rs"),
            ),
        )
        self.assertEqual(
            boundaries.KEYNOTE_SLIDE_DELETE_CANONICAL_TYPES,
            ("Edit", "Patch", "Commit", "Diagnostics", "Error", "LimitKind", "Path"),
        )
        self.assertEqual(
            boundaries.KEYNOTE_SLIDE_DELETE_PACKAGE_METHODS,
            ("edit_slide_deletion", "apply_slide_deletion"),
        )
        self.assertEqual(
            boundaries.KEYNOTE_SLIDE_DELETE_EDIT_METHODS,
            ("remove_slide",),
        )

    def test_retired_iwa_keynote_slide_delete_surface_cannot_return(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            retired = root / boundaries.RETIRED_IWA_KEYNOTE_SLIDE_DELETE_SOURCE
            retired.parent.mkdir(parents=True)
            retired.write_text("pub fn remove_slide() {}\n", encoding="utf-8")
            nested = root / boundaries.IWA_KEYNOTE_SOURCE_ROOT / "legacy/delete.rs"
            nested.parent.mkdir(parents=True)
            nested.write_text(
                "pub(in crate::keynote) async unsafe fn r#remove_slide() {}\n",
                encoding="utf-8",
            )
            editor = root / boundaries.IWA_KEYNOTE_EDITOR_SOURCE
            editor.write_text("pub(crate) mod r#slide_delete;\n", encoding="utf-8")
            example = root / boundaries.RETIRED_IWA_KEYNOTE_SLIDE_DELETE_EXAMPLE
            example.parent.mkdir(parents=True)
            example.write_text("// retired example\n", encoding="utf-8")

            self.assertEqual(
                boundaries.audit_iwa_keynote_slide_delete_source_topology(root),
                sorted(
                    [
                        "retired litchi-iwa Keynote slide-delete example returned: "
                        "crates/litchi-iwa/examples/remove_keynote_slide.rs",
                        "retired litchi-iwa Keynote slide-delete method remove_slide: "
                        "crates/litchi-iwa/src/keynote/editor/slide_delete.rs:1",
                        "retired litchi-iwa Keynote slide-delete method remove_slide: "
                        "crates/litchi-iwa/src/keynote/legacy/delete.rs:1",
                        "retired litchi-iwa Keynote slide-delete module slide_delete: "
                        "crates/litchi-iwa/src/keynote/editor.rs:1",
                        "retired litchi-iwa Keynote slide-delete source returned: "
                        "crates/litchi-iwa/src/keynote/editor/slide_delete.rs",
                    ]
                ),
            )

    def test_retired_iwa_keynote_slide_delete_readme_calls_and_example(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            readme = root / boundaries.IWA_KEYNOTE_README
            readme.parent.mkdir(parents=True)
            readme.write_text(
                "editor.r#remove_slide(0)?;\n"
                "KeynoteEditor::remove_slide(&mut editor, 0)?;\n"
                "Run `remove_keynote_slide` or remove_keynote_slide.rs.\n",
                encoding="utf-8",
            )

            self.assertEqual(
                boundaries.audit_iwa_keynote_slide_delete_source_topology(root),
                [
                    "retired litchi-iwa Keynote slide-delete README call remove_slide: "
                    "crates/litchi-iwa/README.md:1",
                    "retired litchi-iwa Keynote slide-delete README call remove_slide: "
                    "crates/litchi-iwa/README.md:2",
                    "retired litchi-iwa Keynote slide-delete README example reference "
                    "remove_keynote_slide: crates/litchi-iwa/README.md:3",
                ],
            )

    def test_iwa_keynote_slide_delete_policy_ignores_near_names_and_non_code(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            legacy = root / boundaries.IWA_KEYNOTE_SOURCE_ROOT / "legacy/delete.rs"
            legacy.parent.mkdir(parents=True)
            legacy.write_text(
                "// pub fn remove_slide() {}\n"
                'const NOTE: &str = "fn remove_slide() {}";\n'
                "pub fn remove_slides() {}\n"
                "pub fn remove_slide_movie() {}\n",
                encoding="utf-8",
            )
            editor = root / boundaries.IWA_KEYNOTE_EDITOR_SOURCE
            editor.write_text(
                "// mod slide_delete;\nmod slide_deletes;\n",
                encoding="utf-8",
            )
            readme = root / boundaries.IWA_KEYNOTE_README
            readme.write_text(
                "keynote.remove_slides();\n"
                "keynote.remove_slide_movie(0, selector);\n"
                "edit.remove_slide(SlideSelector::name(\"Appendix\"))?;\n"
                "remove_keynote_slides.rs\n",
                encoding="utf-8",
            )

            self.assertEqual(
                boundaries.audit_iwa_keynote_slide_delete_source_topology(root), []
            )

    def test_focused_keynote_slide_delete_requires_exact_topology(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            add_keynote_slide_delete_canonical_scaffold(root)
            self.assertEqual(
                boundaries.audit_keynote_slide_delete_facade_source_topology(root), []
            )

            (root / boundaries.KEYNOTE_SLIDE_DELETE_SEMANTIC_SOURCE).unlink()
            violations = (
                boundaries.audit_keynote_slide_delete_facade_source_topology(root)
            )
            for name in boundaries.KEYNOTE_SLIDE_DELETE_CANONICAL_TYPES:
                self.assertTrue(
                    any(
                        f"canonical slide::delete type {name}" in item
                        for item in violations
                    )
                )
            self.assertTrue(any("Edit::remove_slide" in item for item in violations))

    def test_focused_keynote_slide_delete_rejects_ids_bytes_proto_and_raw_types(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            add_keynote_slide_delete_canonical_scaffold(root)
            semantic = root / boundaries.KEYNOTE_SLIDE_DELETE_SEMANTIC_SOURCE
            semantic.write_text(
                semantic.read_text(encoding="utf-8")
                + "pub fn by_object_id(object_id: u64) -> RawMessage { todo!() }\n"
                + "pub fn from_source_bytes(source_bytes: &[u8]) "
                "-> kn::SlideArchive { todo!() }\n"
                + "pub fn protobuf(value: prost_types::MessageInfo) "
                "-> wire::WireView { todo!() }\n",
                encoding="utf-8",
            )

            violations = (
                boundaries.audit_keynote_slide_delete_facade_source_topology(root)
            )
            for fragment in (
                "raw identifier object_id",
                "raw source bytes source_bytes",
                "raw byte slice &[u8]",
                "archive/IWA type RawMessage",
                "protobuf type kn",
                "archive/IWA type SlideArchive",
                "protobuf type prost_types",
                "archive/IWA type MessageInfo",
                "wire type wire",
                "wire type WireView",
            ):
                with self.subTest(fragment=fragment):
                    self.assertTrue(
                        any(fragment in item for item in violations),
                        msg=f"missing focused slide-delete leak: {fragment}",
                    )

    def test_focused_keynote_slide_delete_rejects_all_flat_and_root_aliases(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            add_keynote_slide_delete_canonical_scaffold(root)
            semantic = root / boundaries.KEYNOTE_SLIDE_DELETE_SEMANTIC_SOURCE
            aliases = sorted(boundaries.KEYNOTE_SLIDE_DELETE_FLAT_ALIASES)
            semantic.write_text(
                semantic.read_text(encoding="utf-8")
                + "".join(f"pub type {alias} = Edit;\n" for alias in aliases),
                encoding="utf-8",
            )
            lib = root / boundaries.KEYNOTE_SLIDE_DELETE_EXPORT_SOURCES[0]
            root_aliases = sorted(boundaries.KEYNOTE_SLIDE_DELETE_SHORT_NAMES)
            lib.write_text(
                lib.read_text(encoding="utf-8")
                + "pub use crate::slide::delete::{"
                + ", ".join(root_aliases)
                + "};\n",
                encoding="utf-8",
            )

            violations = (
                boundaries.audit_keynote_slide_delete_facade_source_topology(root)
            )
            for alias in aliases:
                self.assertTrue(
                    any(f"flat alias {alias}" in item for item in violations)
                )
            for alias in root_aliases:
                self.assertTrue(
                    any(f"root alias {alias}" in item for item in violations)
                )
            self.assertTrue(
                any("public slide-delete owner alias" in item for item in violations)
            )

    def test_focused_keynote_slide_delete_rejects_public_package_module(self) -> None:
        for declaration in (
            "pub mod slide_delete;",
            "pub mod r#slide_delete;",
            "pub mod slide_delete {}",
            "pub\nmod\nr#slide_delete\n{}",
        ):
            with self.subTest(declaration=declaration):
                with tempfile.TemporaryDirectory() as directory:
                    root = Path(directory)
                    add_keynote_slide_delete_canonical_scaffold(root)
                    package = root / boundaries.KEYNOTE_SLIDE_DELETE_EXPORT_SOURCES[1]
                    package.write_text(declaration + "\n", encoding="utf-8")

                    self.assertIn(
                        "focused litchi-keynote slide-delete public API exposes "
                        "duplicate package::slide_delete module: "
                        "crates/litchi-keynote/src/package.rs:1",
                        boundaries.audit_keynote_slide_delete_facade_source_topology(
                            root
                        ),
                    )

    def test_keynote_placeholder_visibility_boundary_inventories_are_exact(
        self,
    ) -> None:
        self.assertEqual(
            boundaries.RETIRED_IWA_KEYNOTE_PLACEHOLDER_VISIBILITY_METHODS,
            (
                "set_slide_text_placeholder_visible",
                "set_slide_title_visible",
                "set_slide_body_visible",
            ),
        )
        self.assertEqual(
            boundaries.RETIRED_IWA_KEYNOTE_PLACEHOLDER_VISIBILITY_SOURCE,
            Path("crates/litchi-iwa/src/keynote/editor/placeholder_visibility.rs"),
        )
        self.assertEqual(
            boundaries.RETIRED_IWA_KEYNOTE_PLACEHOLDER_VISIBILITY_MODULES,
            ("placeholder_visibility",),
        )
        self.assertEqual(
            boundaries.RETIRED_IWA_KEYNOTE_PLACEHOLDER_VISIBILITY_EXAMPLE,
            Path("crates/litchi-iwa/examples/set_keynote_placeholder_visibility.rs"),
        )
        self.assertEqual(
            boundaries.RETIRED_IWA_KEYNOTE_PLACEHOLDER_VISIBILITY_PUBLIC_TYPES,
            ("KeynoteSlideTextPlaceholder",),
        )
        self.assertEqual(
            boundaries.KEYNOTE_PLACEHOLDER_VISIBILITY_OWNER_HELPER_ROOT,
            Path(
                "crates/litchi-keynote/src/package/slide_placeholder_visibility"
            ),
        )
        self.assertEqual(
            boundaries.KEYNOTE_PLACEHOLDER_VISIBILITY_OWNER_HELPER_SOURCES,
            (
                Path(
                    "crates/litchi-keynote/src/package/"
                    "slide_placeholder_visibility/errors.rs"
                ),
                Path(
                    "crates/litchi-keynote/src/package/"
                    "slide_placeholder_visibility/resolve.rs"
                ),
                Path(
                    "crates/litchi-keynote/src/package/"
                    "slide_placeholder_visibility/rewrite.rs"
                ),
                Path(
                    "crates/litchi-keynote/src/package/"
                    "slide_placeholder_visibility/slide_number.rs"
                ),
                Path(
                    "crates/litchi-keynote/src/package/"
                    "slide_placeholder_visibility/verification.rs"
                ),
            ),
        )
        self.assertEqual(
            boundaries.KEYNOTE_PLACEHOLDER_VISIBILITY_SLIDE_PREVIEW_SOURCE,
            Path("crates/litchi-keynote/src/package/slide_preview.rs"),
        )
        self.assertEqual(
            boundaries.KEYNOTE_PLACEHOLDER_VISIBILITY_PREVIEW_HELPER_ROOT,
            Path("crates/litchi-keynote/src/package/slide_preview"),
        )
        self.assertEqual(
            boundaries.KEYNOTE_PLACEHOLDER_VISIBILITY_PREVIEW_HELPER_SOURCES,
            (
                Path(
                    "crates/litchi-keynote/src/package/slide_preview/"
                    "slide_number.rs"
                ),
            ),
        )
        self.assertEqual(
            boundaries.KEYNOTE_PLACEHOLDER_VISIBILITY_IMPLEMENTATION_SOURCES,
            (
                Path("crates/litchi-keynote/src/slide/placeholder.rs"),
                Path(
                    "crates/litchi-keynote/src/package/"
                    "slide_placeholder_visibility.rs"
                ),
                *boundaries.KEYNOTE_PLACEHOLDER_VISIBILITY_OWNER_HELPER_SOURCES,
                *boundaries.KEYNOTE_PLACEHOLDER_VISIBILITY_PREVIEW_HELPER_SOURCES,
            ),
        )
        self.assertEqual(
            boundaries.KEYNOTE_PLACEHOLDER_VISIBILITY_EXPORT_SOURCES,
            (
                Path("crates/litchi-keynote/src/lib.rs"),
                Path("crates/litchi-keynote/src/package.rs"),
                Path("crates/litchi-keynote/src/slide.rs"),
            ),
        )
        self.assertEqual(
            boundaries.KEYNOTE_PLACEHOLDER_VISIBILITY_CANONICAL_TYPES,
            ("Kind", "State", "Edit", "Patch", "Commit", "Diagnostics", "Error", "LimitKind"),
        )
        self.assertEqual(
            boundaries.KEYNOTE_PLACEHOLDER_VISIBILITY_CANONICAL_KINDS,
            ("Title", "Body", "SlideNumber"),
        )
        self.assertEqual(
            boundaries.KEYNOTE_PLACEHOLDER_VISIBILITY_PACKAGE_METHODS,
            (
                "slide_placeholder_visibility",
                "edit_slide_placeholder_visibility",
                "apply_slide_placeholder_visibility",
            ),
        )
        self.assertEqual(
            boundaries.KEYNOTE_PLACEHOLDER_VISIBILITY_FLAT_ALIAS_PREFIXES,
            (
                "Placeholder",
                "PlaceholderVisibility",
                "SlidePlaceholder",
                "SlidePlaceholderVisibility",
                "SlideTextPlaceholder",
                "SlideNumber",
                "SlideNumberPlaceholder",
                "SlideNumberVisibility",
                "SlideNumberPlaceholderVisibility",
            ),
        )
        self.assertEqual(
            boundaries.KEYNOTE_PLACEHOLDER_VISIBILITY_FLAT_ALIASES,
            frozenset(
                prefix + suffix
                for prefix in boundaries.KEYNOTE_PLACEHOLDER_VISIBILITY_FLAT_ALIAS_PREFIXES
                for suffix in boundaries.KEYNOTE_PLACEHOLDER_VISIBILITY_CANONICAL_TYPES
            ),
        )
        self.assertEqual(len(boundaries.KEYNOTE_PLACEHOLDER_VISIBILITY_FLAT_ALIASES), 72)
        self.assertEqual(
            boundaries.KEYNOTE_PLACEHOLDER_VISIBILITY_FLAT_SEMANTIC_ALIASES,
            frozenset({"SlideNumberPlaceholder", "SlideNumberVisibility"}),
        )
        self.assertEqual(
            boundaries.KEYNOTE_PLACEHOLDER_VISIBILITY_SLIDE_NUMBER_PUBLIC_MEMBERS,
            frozenset(
                {
                    "apply_slide_number_visibility",
                    "edit_slide_number_visibility",
                    "hide_slide_number",
                    "is_slide_number_visible",
                    "set_slide_number_visible",
                    "show_slide_number",
                    "slide_number_visibility",
                }
            ),
        )
        self.assertEqual(
            boundaries.KEYNOTE_PLACEHOLDER_VISIBILITY_PHYSICAL_TYPES,
            frozenset(
                {
                    "Archive",
                    "ArchiveObject",
                    "ComponentCatalog",
                    "EntryEdit",
                    "ExactArtifacts",
                    "IWorkPackage",
                    "NodeVisibilitySnapshot",
                    "PhysicalSource",
                    "PlaceholderOwnerSnapshot",
                    "PlaceholderTextOwnerSnapshot",
                    "PlaceholderVisibilitySnapshot",
                    "RawMessage",
                    "ReferenceSnapshot",
                    "Resolved",
                    "SlideOwnerSnapshot",
                    "SlideNodeSnapshot",
                    "SlideNumberSnapshot",
                    "SlideNumberVisibilitySnapshot",
                    "SnappyStream",
                    "SourceCatalog",
                }
            ),
        )
        self.assertNotIn(
            "PlaceholderKind",
            boundaries.KEYNOTE_PLACEHOLDER_VISIBILITY_PHYSICAL_TYPES,
        )
        self.assertEqual(
            boundaries.KEYNOTE_PLACEHOLDER_VISIBILITY_PROTO_ORIGINS,
            frozenset({"kn", "tsp"}),
        )
        self.assertEqual(
            boundaries.KEYNOTE_PLACEHOLDER_VISIBILITY_WIRE_TYPES,
            frozenset(
                {
                    "DecodeLimitKind",
                    "DecodeOptions",
                    "NestedFieldEdit",
                    "NestedFieldReplacement",
                    "WireDescent",
                    "WireError",
                    "WireLimits",
                    "WireResourceLimit",
                    "WireView",
                }
            ),
        )

    def test_retired_iwa_keynote_placeholder_visibility_surface_cannot_return(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            retired = (
                root / boundaries.RETIRED_IWA_KEYNOTE_PLACEHOLDER_VISIBILITY_SOURCE
            )
            retired.parent.mkdir(parents=True)
            retired.write_text(
                "pub struct KeynoteSlideTextPlaceholder;\n"
                "pub fn r#set_slide_text_placeholder_visible(\n"
                "    value: KeynoteSlideTextPlaceholder,\n"
                ") {}\n"
                "pub fn set_slide_title_visible() {}\n",
                encoding="utf-8",
            )
            nested = root / boundaries.IWA_KEYNOTE_SOURCE_ROOT / "nested.rs"
            nested.write_text(
                "pub use crate::keynote::KeynoteSlideTextPlaceholder;\n"
                "fn set_slide_body_visible() {}\n",
                encoding="utf-8",
            )
            editor = root / boundaries.IWA_KEYNOTE_EDITOR_SOURCE
            editor.write_text("pub mod r#placeholder_visibility;\n", encoding="utf-8")
            example = (
                root / boundaries.RETIRED_IWA_KEYNOTE_PLACEHOLDER_VISIBILITY_EXAMPLE
            )
            example.parent.mkdir(parents=True)
            example.write_text("fn main() {}\n", encoding="utf-8")

            self.assertEqual(
                boundaries.audit_iwa_keynote_placeholder_visibility_source_topology(
                    root
                ),
                sorted(
                    [
                        "retired litchi-iwa Keynote placeholder visibility source "
                        "returned: crates/litchi-iwa/src/keynote/editor/"
                        "placeholder_visibility.rs",
                        "retired litchi-iwa Keynote placeholder visibility example "
                        "returned: crates/litchi-iwa/examples/"
                        "set_keynote_placeholder_visibility.rs",
                        "retired litchi-iwa Keynote placeholder visibility module "
                        "placeholder_visibility: crates/litchi-iwa/src/keynote/"
                        "editor.rs:1",
                        "retired litchi-iwa Keynote placeholder visibility method "
                        "set_slide_text_placeholder_visible: crates/litchi-iwa/src/"
                        "keynote/editor/placeholder_visibility.rs:2",
                        "retired litchi-iwa Keynote placeholder visibility method "
                        "set_slide_title_visible: crates/litchi-iwa/src/keynote/"
                        "editor/placeholder_visibility.rs:5",
                        "retired litchi-iwa Keynote placeholder visibility method "
                        "set_slide_body_visible: crates/litchi-iwa/src/keynote/nested.rs:2",
                        "retired litchi-iwa Keynote placeholder visibility public "
                        "type KeynoteSlideTextPlaceholder: crates/litchi-iwa/src/"
                        "keynote/editor/placeholder_visibility.rs:1",
                        "retired litchi-iwa Keynote placeholder visibility public "
                        "type KeynoteSlideTextPlaceholder: crates/litchi-iwa/src/"
                        "keynote/editor/placeholder_visibility.rs:3",
                        "retired litchi-iwa Keynote placeholder visibility public "
                        "type KeynoteSlideTextPlaceholder: crates/litchi-iwa/src/"
                        "keynote/nested.rs:1",
                    ]
                ),
            )

    def test_retired_iwa_keynote_placeholder_visibility_module_variants(
        self,
    ) -> None:
        declarations = (
            "mod placeholder_visibility;",
            "pub mod r#placeholder_visibility;",
            "pub(crate) mod placeholder_visibility {}",
            "pub(super) mod r#placeholder_visibility {}",
            "pub(in crate) mod placeholder_visibility;",
            "pub\nmod\nr#placeholder_visibility\n{}",
        )
        for declaration in declarations:
            with self.subTest(declaration=declaration):
                with tempfile.TemporaryDirectory() as directory:
                    root = Path(directory)
                    editor = root / boundaries.IWA_KEYNOTE_EDITOR_SOURCE
                    editor.parent.mkdir(parents=True)
                    editor.write_text(declaration + "\n", encoding="utf-8")

                    self.assertEqual(
                        boundaries.audit_iwa_keynote_placeholder_visibility_source_topology(
                            root
                        ),
                        [
                            "retired litchi-iwa Keynote placeholder visibility module "
                            "placeholder_visibility: crates/litchi-iwa/src/keynote/"
                            "editor.rs:1"
                        ],
                    )

    def test_retired_iwa_keynote_placeholder_visibility_readme_calls_and_example(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            readme = root / boundaries.IWA_KEYNOTE_README
            readme.parent.mkdir(parents=True)
            readme.write_text(
                "keynote.r#set_slide_text_placeholder_visible(1);\n"
                "canvas\n    .\n    set_slide_title_visible(true);\n"
                "crate::nested::KeynoteEditor::\n"
                "    r#set_slide_body_visible(false);\n"
                "set_keynote_placeholder_visibility\n"
                "set_keynote_placeholder_visibility.rs\n",
                encoding="utf-8",
            )

            self.assertEqual(
                boundaries.audit_iwa_keynote_placeholder_visibility_source_topology(
                    root
                ),
                sorted(
                    [
                        "retired litchi-iwa Keynote placeholder visibility README "
                        "call set_slide_text_placeholder_visible: crates/litchi-iwa/"
                        "README.md:1",
                        "retired litchi-iwa Keynote placeholder visibility README "
                        "call set_slide_title_visible: crates/litchi-iwa/README.md:4",
                        "retired litchi-iwa Keynote placeholder visibility README "
                        "call set_slide_body_visible: crates/litchi-iwa/README.md:6",
                        "retired litchi-iwa Keynote placeholder visibility README "
                        "example reference set_keynote_placeholder_visibility: "
                        "crates/litchi-iwa/README.md:7",
                        "retired litchi-iwa Keynote placeholder visibility README "
                        "example reference set_keynote_placeholder_visibility: "
                        "crates/litchi-iwa/README.md:8",
                    ]
                ),
            )

    def test_iwa_keynote_placeholder_visibility_policy_ignores_safe_surfaces(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            editor = root / boundaries.IWA_KEYNOTE_EDITOR_SOURCE
            editor.parent.mkdir(parents=True)
            editor.write_text(
                "// mod placeholder_visibility;\n"
                'const NOTE: &str = "pub mod placeholder_visibility;";\n'
                'const RAW: &str = r#"mod r#placeholder_visibility {}"#;\n'
                "/* outer /* mod placeholder_visibility; */ still comment */\n"
                "mod placeholder_visibilities;\n",
                encoding="utf-8",
            )
            safe = root / boundaries.IWA_KEYNOTE_SOURCE_ROOT / "safe.rs"
            safe.write_text(
                "// fn set_slide_title_visible() {}\n"
                'const FN_NOTE: &str = "fn set_slide_body_visible() {}";\n'
                "/* fn r#set_slide_text_placeholder_visible() {} */\n"
                "pub(crate) struct KeynoteSlideTextPlaceholder;\n"
                "pub(super) use crate::KeynoteSlideTextPlaceholder;\n"
                "pub(in crate) fn retained(value: KeynoteSlideTextPlaceholder) {}\n"
                "pub struct KeynoteSlideTextPlaceholders;\n"
                "pub fn set_placeholder_visible() {}\n"
                "pub fn slide_layout() {}\n"
                "pub fn set_slide_layout() {}\n"
                "pub fn slide_number_visible() {}\n"
                "pub fn set_slide_number_visible() {}\n"
                "pub struct KeynoteSlideInfo;\n",
                encoding="utf-8",
            )
            other_owner = root / "crates/litchi-pages/src/editor.rs"
            other_owner.parent.mkdir(parents=True)
            other_owner.write_text(
                "pub struct KeynoteSlideTextPlaceholder;\n"
                "pub fn set_slide_title_visible() {}\n"
                "pub fn set_slide_body_visible() {}\n"
                "pub fn set_slide_text_placeholder_visible() {}\n",
                encoding="utf-8",
            )
            readme = root / boundaries.IWA_KEYNOTE_README
            readme.parent.mkdir(parents=True, exist_ok=True)
            readme.write_text(
                "`set_slide_title_visible` is retired prose.\n"
                "set_slide_body_visible\n"
                "Package::slide_placeholder_visibility()\n"
                "package.edit_slide_placeholder_visibility()\n"
                "package.apply_slide_placeholder_visibility()\n"
                "package.show_slide_title()\n"
                "package.hide_slide_body()\n"
                "set_keynote_placeholder_visibilities.rs\n",
                encoding="utf-8",
            )
            other_readme = root / "README.md"
            other_readme.write_text(
                "editor.set_slide_title_visible(true);\n"
                "set_keynote_placeholder_visibility.rs\n",
                encoding="utf-8",
            )

            self.assertEqual(
                boundaries.audit_iwa_keynote_placeholder_visibility_source_topology(
                    root
                ),
                [],
            )

    def test_iwa_keynote_slide_number_visibility_retirement_inventory_is_exact(
        self,
    ) -> None:
        self.assertEqual(
            boundaries.RETIRED_IWA_KEYNOTE_SLIDE_NUMBER_VISIBILITY_METHODS,
            ("set_slide_number_visible",),
        )
        self.assertEqual(
            boundaries.RETIRED_IWA_KEYNOTE_SLIDE_NUMBER_VISIBILITY_SOURCE,
            Path("crates/litchi-iwa/src/keynote/editor/slide_number.rs"),
        )
        self.assertEqual(
            boundaries.RETIRED_IWA_KEYNOTE_SLIDE_NUMBER_VISIBILITY_MODULES,
            ("slide_number",),
        )
        self.assertEqual(
            boundaries.RETIRED_IWA_KEYNOTE_SLIDE_NUMBER_VISIBILITY_EXAMPLE,
            Path("crates/litchi-iwa/examples/set_keynote_slide_number.rs"),
        )
        self.assertEqual(
            boundaries.RETIRED_IWA_KEYNOTE_SLIDE_NUMBER_VISIBILITY_TESTS,
            (
                "slide_number_visibility_matches_native_ownership_and_round_trips_exactly",
                "slide_number_visibility_rejects_inconsistent_native_state_transactionally",
            ),
        )

    def test_retired_iwa_keynote_slide_number_visibility_surface_cannot_return(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            retired = (
                root
                / boundaries.RETIRED_IWA_KEYNOTE_SLIDE_NUMBER_VISIBILITY_SOURCE
            )
            retired.parent.mkdir(parents=True)
            retired.write_text(
                "pub fn r#set_slide_number_visible() {}\n", encoding="utf-8"
            )
            editor = root / boundaries.IWA_KEYNOTE_EDITOR_SOURCE
            editor.write_text("pub mod r#slide_number;\n", encoding="utf-8")
            tests = root / boundaries.IWA_KEYNOTE_EDITOR_TEST_SOURCE
            tests.write_text(
                "fn slide_number_visibility_matches_native_ownership_and_round_trips_exactly() {}\n"
                "fn slide_number_visibility_rejects_inconsistent_native_state_transactionally() {}\n",
                encoding="utf-8",
            )
            example = (
                root / boundaries.RETIRED_IWA_KEYNOTE_SLIDE_NUMBER_VISIBILITY_EXAMPLE
            )
            example.parent.mkdir(parents=True)
            example.write_text("fn main() {}\n", encoding="utf-8")

            self.assertEqual(
                boundaries.audit_iwa_keynote_slide_number_visibility_source_topology(
                    root
                ),
                sorted(
                    [
                        "retired litchi-iwa Keynote slide-number visibility source "
                        "returned: crates/litchi-iwa/src/keynote/editor/slide_number.rs",
                        "retired litchi-iwa Keynote slide-number visibility example "
                        "returned: crates/litchi-iwa/examples/"
                        "set_keynote_slide_number.rs",
                        "retired litchi-iwa Keynote slide-number visibility method "
                        "set_slide_number_visible: crates/litchi-iwa/src/keynote/"
                        "editor/slide_number.rs:1",
                        "retired litchi-iwa Keynote slide-number visibility module "
                        "slide_number: crates/litchi-iwa/src/keynote/editor.rs:1",
                        "retired litchi-iwa Keynote slide-number visibility test "
                        "slide_number_visibility_matches_native_ownership_and_round_"
                        "trips_exactly: crates/litchi-iwa/src/keynote/editor/tests.rs:1",
                        "retired litchi-iwa Keynote slide-number visibility test "
                        "slide_number_visibility_rejects_inconsistent_native_state_"
                        "transactionally: crates/litchi-iwa/src/keynote/editor/tests.rs:2",
                    ]
                ),
            )

    def test_retired_iwa_keynote_slide_number_visibility_module_variants(
        self,
    ) -> None:
        for declaration in (
            "mod slide_number;",
            "pub mod r#slide_number;",
            "pub(crate) mod slide_number {}",
            "pub(super) mod r#slide_number {}",
            "pub(in crate) mod slide_number;",
            "pub\nmod\nr#slide_number\n{}",
        ):
            with self.subTest(declaration=declaration):
                with tempfile.TemporaryDirectory() as directory:
                    root = Path(directory)
                    editor = root / boundaries.IWA_KEYNOTE_EDITOR_SOURCE
                    editor.parent.mkdir(parents=True)
                    editor.write_text(declaration + "\n", encoding="utf-8")

                    self.assertEqual(
                        boundaries.audit_iwa_keynote_slide_number_visibility_source_topology(
                            root
                        ),
                        [
                            "retired litchi-iwa Keynote slide-number visibility "
                            "module slide_number: crates/litchi-iwa/src/keynote/"
                            "editor.rs:1"
                        ],
                    )

    def test_retired_iwa_keynote_slide_number_visibility_readme_calls_and_example(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            readme = root / boundaries.IWA_KEYNOTE_README
            readme.parent.mkdir(parents=True)
            readme.write_text(
                "editor.r#set_slide_number_visible(true);\n"
                "canvas\n    .\n    set_slide_number_visible(false);\n"
                "crate::nested::KeynoteEditor::\n"
                "    r#set_slide_number_visible(true);\n"
                "set_keynote_slide_number\n"
                "set_keynote_slide_number.rs\n",
                encoding="utf-8",
            )

            self.assertEqual(
                boundaries.audit_iwa_keynote_slide_number_visibility_source_topology(
                    root
                ),
                sorted(
                    [
                        "retired litchi-iwa Keynote slide-number visibility README "
                        "call set_slide_number_visible: crates/litchi-iwa/README.md:1",
                        "retired litchi-iwa Keynote slide-number visibility README "
                        "call set_slide_number_visible: crates/litchi-iwa/README.md:4",
                        "retired litchi-iwa Keynote slide-number visibility README "
                        "call set_slide_number_visible: crates/litchi-iwa/README.md:6",
                        "retired litchi-iwa Keynote slide-number visibility README "
                        "example reference set_keynote_slide_number: "
                        "crates/litchi-iwa/README.md:7",
                        "retired litchi-iwa Keynote slide-number visibility README "
                        "example reference set_keynote_slide_number: "
                        "crates/litchi-iwa/README.md:8",
                    ]
                ),
            )

    def test_iwa_keynote_slide_number_visibility_policy_retains_shared_surfaces(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            editor = root / boundaries.IWA_KEYNOTE_EDITOR_SOURCE
            editor.parent.mkdir(parents=True)
            editor.write_text(
                "// mod slide_number;\n"
                'const NOTE: &str = "pub mod slide_number;";\n'
                'const RAW: &str = r#"mod r#slide_number {}"#;\n'
                "/* outer /* mod slide_number; */ still comment */\n"
                "mod slide_numbers;\n"
                "pub struct KeynoteSlideInfo {\n"
                "    pub is_slide_number_visible: Option<bool>,\n"
                "}\n"
                "pub struct KeynoteSlideLayout {\n"
                "    pub slide_number_placeholder_visibility: Option<bool>,\n"
                "}\n",
                encoding="utf-8",
            )
            safe = root / boundaries.IWA_KEYNOTE_SOURCE_ROOT / "creation.rs"
            safe.write_text(
                "// fn set_slide_number_visible() {}\n"
                'const FN_NOTE: &str = "fn set_slide_number_visible() {}";\n'
                "/* fn r#set_slide_number_visible() {} */\n"
                "pub const fn slide_number_visible(self, visible: bool) -> Self "
                "{ self }\n"
                "pub fn set_slide_numbers_visible() {}\n"
                "pub fn set_slide_number_visibility() {}\n",
                encoding="utf-8",
            )
            tests = root / boundaries.IWA_KEYNOTE_EDITOR_TEST_SOURCE
            tests.parent.mkdir(parents=True, exist_ok=True)
            tests.write_text(
                "// fn slide_number_visibility_matches_native_ownership_and_round_trips_exactly() {}\n"
                'const TEST_NOTE: &str = "fn slide_number_visibility_rejects_inconsistent_native_state_transactionally() {}";\n'
                "fn slide_number_visibility_matches_native_ownership_and_round_trips_exactly_v2() {}\n",
                encoding="utf-8",
            )
            other_owner = root / "crates/litchi-pages/src/editor.rs"
            other_owner.parent.mkdir(parents=True)
            other_owner.write_text(
                "pub fn set_slide_number_visible() {}\n", encoding="utf-8"
            )
            retained_example = (
                root / "crates/litchi-iwa/examples/create_keynote_slide_numbers.rs"
            )
            retained_example.parent.mkdir(parents=True)
            retained_example.write_text("fn main() {}\n", encoding="utf-8")
            readme = root / boundaries.IWA_KEYNOTE_README
            readme.parent.mkdir(parents=True, exist_ok=True)
            readme.write_text(
                "`set_slide_number_visible` is retired prose.\n"
                "set_slide_number_visible\n"
                "package.slide_placeholder_visibility(slide, Kind::SlideNumber)\n"
                "package.edit_slide_placeholder_visibility(slide, Kind::SlideNumber)\n"
                "settings.set_slide_numbers_visible(true)\n"
                "create_keynote_slide_numbers.rs\n"
                "edit_slide_number_visibility.rs\n",
                encoding="utf-8",
            )
            other_readme = root / "README.md"
            other_readme.write_text(
                "editor.set_slide_number_visible(true)\n"
                "set_keynote_slide_number.rs\n",
                encoding="utf-8",
            )

            self.assertEqual(
                boundaries.audit_iwa_keynote_slide_number_visibility_source_topology(
                    root
                ),
                [],
            )

    def test_focused_keynote_placeholder_visibility_requires_each_canonical_type(
        self,
    ) -> None:
        for missing in boundaries.KEYNOTE_PLACEHOLDER_VISIBILITY_CANONICAL_TYPES:
            with self.subTest(missing=missing):
                with tempfile.TemporaryDirectory() as directory:
                    root = Path(directory)
                    add_keynote_placeholder_visibility_canonical_scaffold(root)
                    semantic = (
                        root
                        / boundaries.KEYNOTE_PLACEHOLDER_VISIBILITY_SEMANTIC_SOURCE
                    )
                    semantic.write_text(
                        (
                            "pub enum Kind { Title, Body, SlideNumber }\n"
                            if missing != "Kind"
                            else ""
                        )
                        + "".join(
                            f"pub struct {name};\n"
                            for name in boundaries.KEYNOTE_PLACEHOLDER_VISIBILITY_CANONICAL_TYPES
                            if name not in {"Kind", missing}
                        ),
                        encoding="utf-8",
                    )

                    expected = [
                        "focused litchi-keynote placeholder visibility public "
                        "API is missing canonical slide::placeholder type "
                        f"{missing}: crates/litchi-keynote/src/slide/placeholder.rs"
                    ]
                    self.assertEqual(
                        boundaries.audit_keynote_placeholder_visibility_facade_source_topology(
                            root
                        ),
                        sorted(expected),
                    )

    def test_focused_keynote_placeholder_visibility_requires_each_canonical_kind(
        self,
    ) -> None:
        for missing in boundaries.KEYNOTE_PLACEHOLDER_VISIBILITY_CANONICAL_KINDS:
            with self.subTest(missing=missing):
                with tempfile.TemporaryDirectory() as directory:
                    root = Path(directory)
                    add_keynote_placeholder_visibility_canonical_scaffold(root)
                    semantic = (
                        root
                        / boundaries.KEYNOTE_PLACEHOLDER_VISIBILITY_SEMANTIC_SOURCE
                    )
                    retained = [
                        kind
                        for kind in boundaries.KEYNOTE_PLACEHOLDER_VISIBILITY_CANONICAL_KINDS
                        if kind != missing
                    ]
                    semantic.write_text(
                        f"// {missing},\n"
                        f'const NOTE: &str = "{missing}";\n'
                        "pub enum Kind { "
                        + ", ".join(retained)
                        + " }\n"
                        + "".join(
                            f"pub struct {name};\n"
                            for name in boundaries.KEYNOTE_PLACEHOLDER_VISIBILITY_CANONICAL_TYPES
                            if name != "Kind"
                        ),
                        encoding="utf-8",
                    )

                    self.assertEqual(
                        boundaries.audit_keynote_placeholder_visibility_facade_source_topology(
                            root
                        ),
                        [
                            "focused litchi-keynote placeholder visibility public "
                            "API is missing canonical placeholder kind "
                            f"{missing}: crates/litchi-keynote/src/slide/placeholder.rs"
                        ],
                    )

    def test_focused_keynote_placeholder_visibility_requires_nested_modules_and_private_owner(
        self,
    ) -> None:
        missing_root = (
            "",
            "mod slide;\n",
            "pub(crate) mod slide;\n",
            "pub(super) mod r#slide {}\n",
            "pub(in crate) mod slide;\n",
            "// pub mod slide;\n",
            'const NOTE: &str = "pub mod slide;";\n',
        )
        accepted_root = (
            "pub mod slide;\n",
            "pub mod r#slide;\n",
            "pub mod slide {}\n",
            "pub\nmod\nr#slide\n{}\n",
        )
        for declaration, expected in (
            *[
                (
                    declaration,
                    [
                        "focused litchi-keynote placeholder visibility public API "
                        "is missing canonical root slide module: "
                        "crates/litchi-keynote/src/lib.rs"
                    ],
                )
                for declaration in missing_root
            ],
            *[(declaration, []) for declaration in accepted_root],
        ):
            with self.subTest(scope="root", declaration=declaration):
                with tempfile.TemporaryDirectory() as directory:
                    root = Path(directory)
                    add_keynote_placeholder_visibility_canonical_scaffold(root)
                    lib_export = (
                        root
                        / boundaries.KEYNOTE_PLACEHOLDER_VISIBILITY_EXPORT_SOURCES[0]
                    )
                    lib_export.write_text(declaration, encoding="utf-8")
                    self.assertEqual(
                        boundaries.audit_keynote_placeholder_visibility_facade_source_topology(
                            root
                        ),
                        expected,
                    )

        missing_placeholder = (
            "",
            "mod placeholder;\n",
            "pub(crate) mod placeholder;\n",
            "pub(super) mod r#placeholder {}\n",
            "pub(in crate) mod placeholder;\n",
            "// pub mod placeholder;\n",
            'const NOTE: &str = "pub mod placeholder;";\n',
        )
        accepted_placeholder = (
            "pub mod placeholder;\n",
            "pub mod r#placeholder;\n",
            "pub mod placeholder {}\n",
            "pub\nmod\nr#placeholder\n{}\n",
        )
        for declaration, expected in (
            *[
                (
                    declaration,
                    [
                        "focused litchi-keynote placeholder visibility public API "
                        "is missing canonical slide::placeholder module: "
                        "crates/litchi-keynote/src/slide.rs"
                    ],
                )
                for declaration in missing_placeholder
            ],
            *[(declaration, []) for declaration in accepted_placeholder],
        ):
            with self.subTest(scope="placeholder", declaration=declaration):
                with tempfile.TemporaryDirectory() as directory:
                    root = Path(directory)
                    add_keynote_placeholder_visibility_canonical_scaffold(root)
                    slide_export = (
                        root
                        / boundaries.KEYNOTE_PLACEHOLDER_VISIBILITY_EXPORT_SOURCES[2]
                    )
                    slide_export.write_text(declaration, encoding="utf-8")
                    self.assertEqual(
                        boundaries.audit_keynote_placeholder_visibility_facade_source_topology(
                            root
                        ),
                        expected,
                    )

        missing_owner_module = (
            "",
            "// mod slide_placeholder_visibility;\n",
            'const NOTE: &str = "mod slide_placeholder_visibility;";\n',
            "mod slide_placeholder_visibilities;\n",
        )
        accepted_owner_module = (
            "mod slide_placeholder_visibility;\n",
            "mod r#slide_placeholder_visibility {}\n",
            "pub(crate) mod slide_placeholder_visibility;\n",
            "pub(super) mod r#slide_placeholder_visibility {}\n",
            "pub(in crate)\nmod\nslide_placeholder_visibility\n;\n",
        )
        for declaration, expected in (
            *[
                (
                    declaration,
                    [
                        "focused litchi-keynote placeholder visibility public API "
                        "is missing private package owner module: "
                        "crates/litchi-keynote/src/package.rs"
                    ],
                )
                for declaration in missing_owner_module
            ],
            *[(declaration, []) for declaration in accepted_owner_module],
        ):
            with self.subTest(scope="owner", declaration=declaration):
                with tempfile.TemporaryDirectory() as directory:
                    root = Path(directory)
                    add_keynote_placeholder_visibility_canonical_scaffold(root)
                    package_export = (
                        root
                        / boundaries.KEYNOTE_PLACEHOLDER_VISIBILITY_EXPORT_SOURCES[1]
                    )
                    package_export.write_text(declaration, encoding="utf-8")
                    self.assertEqual(
                        boundaries.audit_keynote_placeholder_visibility_facade_source_topology(
                            root
                        ),
                        expected,
                    )

        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            add_keynote_placeholder_visibility_canonical_scaffold(root)
            (root / boundaries.KEYNOTE_PLACEHOLDER_VISIBILITY_OWNER_SOURCE).unlink()
            self.assertEqual(
                boundaries.audit_keynote_placeholder_visibility_facade_source_topology(
                    root
                ),
                [
                    "focused litchi-keynote placeholder visibility public API is "
                    "missing private package owner source: crates/litchi-keynote/"
                    "src/package/slide_placeholder_visibility.rs"
                ],
            )

    def test_focused_keynote_placeholder_visibility_rejects_number_specific_aliases_and_members(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            relative_sources = (
                boundaries.KEYNOTE_PLACEHOLDER_VISIBILITY_IMPLEMENTATION_SOURCES
                + boundaries.KEYNOTE_PLACEHOLDER_VISIBILITY_EXPORT_SOURCES
            )
            sources = tuple(root / path for path in relative_sources)
            declarations: dict[Path, list[str]] = {path: [] for path in sources}
            expected: list[str] = []
            semantic_aliases = sorted(
                boundaries.KEYNOTE_PLACEHOLDER_VISIBILITY_FLAT_SEMANTIC_ALIASES
            )
            members = sorted(
                boundaries.KEYNOTE_PLACEHOLDER_VISIBILITY_SLIDE_NUMBER_PUBLIC_MEMBERS
            )
            for index, name in enumerate(semantic_aliases):
                source_index = index % len(sources)
                path = sources[source_index]
                declarations[path].append(f"pub struct {name};")
                expected.append(
                    "focused litchi-keynote placeholder visibility public API "
                    f"retains flat semantic alias {name}: "
                    f"{relative_sources[source_index]}:{len(declarations[path])}"
                )
            for index, name in enumerate(members):
                source_index = (index + len(semantic_aliases)) % len(sources)
                path = sources[source_index]
                declarations[path].append(f"pub fn {name}() {{}}")
                expected.append(
                    "focused litchi-keynote placeholder visibility public API "
                    f"retains slide-number-specific public member {name}: "
                    f"{relative_sources[source_index]}:{len(declarations[path])}"
                )
            for path, lines in declarations.items():
                path.parent.mkdir(parents=True, exist_ok=True)
                path.write_text("\n".join(lines) + "\n", encoding="utf-8")
            add_keynote_placeholder_visibility_canonical_scaffold(root)

            self.assertEqual(
                boundaries.audit_keynote_placeholder_visibility_facade_source_topology(
                    root
                ),
                sorted(expected),
            )

    def test_focused_keynote_placeholder_visibility_rejects_public_number_helper_modules(
        self,
    ) -> None:
        module_sources = (
            boundaries.KEYNOTE_PLACEHOLDER_VISIBILITY_SEMANTIC_SOURCE,
            boundaries.KEYNOTE_PLACEHOLDER_VISIBILITY_OWNER_SOURCE,
            boundaries.KEYNOTE_PLACEHOLDER_VISIBILITY_SLIDE_PREVIEW_SOURCE,
            *boundaries.KEYNOTE_PLACEHOLDER_VISIBILITY_EXPORT_SOURCES,
        )
        declarations = (
            "pub mod slide_number;",
            "pub mod r#slide_number;",
            "pub mod slide_number {}",
            "pub\nmod\nr#slide_number\n{}",
        )
        for relative in module_sources:
            for declaration in declarations:
                with self.subTest(relative=relative, declaration=declaration):
                    with tempfile.TemporaryDirectory() as directory:
                        root = Path(directory)
                        add_keynote_placeholder_visibility_canonical_scaffold(root)
                        path = root / relative
                        path.parent.mkdir(parents=True, exist_ok=True)
                        source = path.read_text(encoding="utf-8") if path.is_file() else ""
                        line = source.count("\n") + 1
                        path.write_text(source + declaration + "\n", encoding="utf-8")

                        self.assertEqual(
                            boundaries.audit_keynote_placeholder_visibility_facade_source_topology(
                                root
                            ),
                            [
                                "focused litchi-keynote placeholder visibility public "
                                "API exposes public slide-number helper module: "
                                f"{relative}:{line}"
                            ],
                        )

    def test_focused_keynote_placeholder_visibility_rejects_public_number_owner_aliases(
        self,
    ) -> None:
        alias_sources = (
            boundaries.KEYNOTE_PLACEHOLDER_VISIBILITY_SEMANTIC_SOURCE,
            *boundaries.KEYNOTE_PLACEHOLDER_VISIBILITY_EXPORT_SOURCES,
        )
        declarations = (
            "pub use crate::slide_number as legacy;",
            "pub use crate::r#slide_number as legacy;",
            "pub\nuse\ncrate::slide_number as legacy;",
            "pub type slide_number = bool;",
        )
        for relative in alias_sources:
            for declaration in declarations:
                with self.subTest(relative=relative, declaration=declaration):
                    with tempfile.TemporaryDirectory() as directory:
                        root = Path(directory)
                        add_keynote_placeholder_visibility_canonical_scaffold(root)
                        path = root / relative
                        source = path.read_text(encoding="utf-8")
                        line = source.count("\n") + 1
                        path.write_text(source + declaration + "\n", encoding="utf-8")

                        self.assertEqual(
                            boundaries.audit_keynote_placeholder_visibility_facade_source_topology(
                                root
                            ),
                            [
                                "focused litchi-keynote placeholder visibility public "
                                "API exposes public slide-number owner alias: "
                                f"{relative}:{line}"
                            ],
                        )

    def test_focused_keynote_placeholder_visibility_rejects_all_flat_aliases(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            relative_sources = (
                boundaries.KEYNOTE_PLACEHOLDER_VISIBILITY_IMPLEMENTATION_SOURCES
                + boundaries.KEYNOTE_PLACEHOLDER_VISIBILITY_EXPORT_SOURCES
            )
            sources = tuple(root / path for path in relative_sources)
            declarations: dict[Path, list[str]] = {path: [] for path in sources}
            expected: list[str] = []
            for index, name in enumerate(
                sorted(boundaries.KEYNOTE_PLACEHOLDER_VISIBILITY_FLAT_ALIASES)
            ):
                source_index = index % len(sources)
                path = sources[source_index]
                if source_index == 0:
                    declaration = f"pub struct {name};"
                elif source_index == 1:
                    declaration = f"pub type {name} = bool;"
                else:
                    declaration = f"pub use crate::legacy::Legacy as {name};"
                declarations[path].append(declaration)
                expected.append(
                    "focused litchi-keynote placeholder visibility public API "
                    f"retains flat alias {name}: {relative_sources[source_index]}:"
                    f"{len(declarations[path])}"
                )
            for path, lines in declarations.items():
                path.parent.mkdir(parents=True, exist_ok=True)
                path.write_text("\n".join(lines) + "\n", encoding="utf-8")
            add_keynote_placeholder_visibility_canonical_scaffold(root)

            self.assertEqual(
                boundaries.audit_keynote_placeholder_visibility_facade_source_topology(
                    root
                ),
                sorted(expected),
            )

    def test_focused_keynote_placeholder_visibility_exports_reject_root_aliases_and_glob(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            exports = tuple(
                root / path
                for path in boundaries.KEYNOTE_PLACEHOLDER_VISIBILITY_EXPORT_SOURCES
            )
            relative = boundaries.KEYNOTE_PLACEHOLDER_VISIBILITY_EXPORT_SOURCES
            aliases = sorted(boundaries.KEYNOTE_PLACEHOLDER_VISIBILITY_SHORT_NAMES)
            grouped = [aliases[0:3], aliases[3:6], aliases[6:]]
            expected: list[str] = []
            for index, (path, names) in enumerate(zip(exports, grouped)):
                path.parent.mkdir(parents=True, exist_ok=True)
                owner = "placeholder" if index == 2 else "slide::placeholder"
                path.write_text(
                    f"pub use crate::{owner}::{{{', '.join(names)}}};\n"
                    f"pub use crate::{owner}::*;\n",
                    encoding="utf-8",
                )
                expected.extend(
                    "focused litchi-keynote placeholder visibility public API "
                    f"retains root alias {name}: {relative[index]}:1"
                    for name in names
                )
                expected.append(
                    "focused litchi-keynote placeholder visibility public API "
                    "retains root aliases via slide::placeholder glob: "
                    f"{relative[index]}:2"
                )
            add_keynote_placeholder_visibility_canonical_scaffold(root)

            self.assertEqual(
                boundaries.audit_keynote_placeholder_visibility_facade_source_topology(
                    root
                ),
                sorted(expected),
            )

    def test_focused_keynote_placeholder_visibility_rejects_duplicate_package_module(
        self,
    ) -> None:
        for declaration in (
            "pub mod slide_placeholder_visibility;",
            "pub mod r#slide_placeholder_visibility;",
            "pub mod slide_placeholder_visibility {}",
            "pub\nmod\nr#slide_placeholder_visibility\n{}",
        ):
            with self.subTest(declaration=declaration):
                with tempfile.TemporaryDirectory() as directory:
                    root = Path(directory)
                    add_keynote_placeholder_visibility_canonical_scaffold(root)
                    package_export = (
                        root
                        / boundaries.KEYNOTE_PLACEHOLDER_VISIBILITY_EXPORT_SOURCES[1]
                    )
                    package_export.write_text(declaration + "\n", encoding="utf-8")

                    self.assertEqual(
                        boundaries.audit_keynote_placeholder_visibility_facade_source_topology(
                            root
                        ),
                        [
                            "focused litchi-keynote placeholder visibility public "
                            "API exposes duplicate package::slide_placeholder_visibility "
                            "module: crates/litchi-keynote/src/package.rs:1"
                        ],
                    )

    def test_focused_keynote_placeholder_visibility_rejects_physical_leaks(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            semantic = (
                root / boundaries.KEYNOTE_PLACEHOLDER_VISIBILITY_SEMANTIC_SOURCE
            )
            semantic.parent.mkdir(parents=True)
            semantic.write_text(
                "pub fn visibility(r#source_bytes: &[u8], r#object_id: u64) "
                "-> (DocumentArchive, IWorkPackage, SourceCatalog) { todo!() }\n"
                "pub type EffectPayload = buffa::DocumentArchiveView;\n"
                "impl prost::Message for placeholder::State {}\n",
                encoding="utf-8",
            )
            owner = root / boundaries.KEYNOTE_PLACEHOLDER_VISIBILITY_OWNER_SOURCE
            owner.parent.mkdir(parents=True, exist_ok=True)
            owner_lines = [
                f"pub type Physical{index} = {name};"
                for index, name in enumerate(
                    sorted(boundaries.KEYNOTE_PLACEHOLDER_VISIBILITY_PHYSICAL_TYPES)
                )
            ]
            owner_lines.extend(
                f"pub type WireLeak{index} = {name};"
                for index, name in enumerate(
                    sorted(boundaries.KEYNOTE_PLACEHOLDER_VISIBILITY_WIRE_TYPES)
                )
            )
            owner_lines.append(
                "pub type Projected = (kn::SlideArchive, tsp::StorageArchive);"
            )
            owner.write_text("\n".join(owner_lines) + "\n", encoding="utf-8")
            lib_export, package_export, slide_export = (
                root / path
                for path in boundaries.KEYNOTE_PLACEHOLDER_VISIBILITY_EXPORT_SOURCES
            )
            lib_export.write_text(
                "pub fn edit_slide_placeholder_visibility("
                "value: litchi_iwa_protos::GeneratedPlaceholder) "
                "-> placeholder::Patch { todo!() }\n",
                encoding="utf-8",
            )
            package_export.write_text(
                "pub fn apply_slide_placeholder_visibility("
                "value: prost_types::MessageInfo) "
                "-> slide::placeholder::Commit { todo!() }\n",
                encoding="utf-8",
            )
            slide_export.write_text(
                "pub fn slide_placeholder_visibility() -> SourceBytes { todo!() }\n",
                encoding="utf-8",
            )
            add_keynote_placeholder_visibility_canonical_scaffold(root)

            violations = (
                boundaries.audit_keynote_placeholder_visibility_facade_source_topology(
                    root
                )
            )

            self.assertEqual(violations, sorted(violations))
            self.assertTrue(
                all(
                    violation.startswith(
                        "focused litchi-keynote placeholder visibility public API "
                        "exposes "
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
                "archive/IWA type litchi_iwa_protos",
                "generated type GeneratedPlaceholder",
                "protobuf type prost_types",
                "archive/IWA type MessageInfo",
                "raw source bytes SourceBytes",
                "protobuf type kn",
                "archive/IWA type SlideArchive",
                "protobuf type tsp",
                "archive/IWA type StorageArchive",
                *tuple(
                    f"archive/IWA type {name}"
                    for name in sorted(
                        boundaries.KEYNOTE_PLACEHOLDER_VISIBILITY_PHYSICAL_TYPES
                    )
                ),
                *tuple(
                    f"wire type {name}"
                    for name in sorted(
                        boundaries.KEYNOTE_PLACEHOLDER_VISIBILITY_WIRE_TYPES
                    )
                ),
            )
            self.assertEqual(len(violations), 48)
            for fragment in expected_fragments:
                with self.subTest(fragment=fragment):
                    self.assertTrue(
                        any(fragment in violation for violation in violations),
                        msg=f"missing focused placeholder-visibility leak: {fragment}",
                    )

    def test_focused_keynote_placeholder_visibility_scans_every_private_helper(
        self,
    ) -> None:
        helpers = (
            boundaries.KEYNOTE_PLACEHOLDER_VISIBILITY_OWNER_HELPER_SOURCES
            + boundaries.KEYNOTE_PLACEHOLDER_VISIBILITY_PREVIEW_HELPER_SOURCES
        )
        for relative in helpers:
            with self.subTest(relative=relative):
                with tempfile.TemporaryDirectory() as directory:
                    root = Path(directory)
                    add_keynote_placeholder_visibility_canonical_scaffold(root)
                    helper = root / relative
                    helper.parent.mkdir(parents=True, exist_ok=True)
                    helper.write_text(
                        "pub fn expose(r#source_bytes: &[u8], r#object_id: u64) "
                        "-> ArchiveObject { todo!() }\n"
                        "pub type PlaceholderEdit = WireView;\n"
                        "impl prost::Message for PlaceholderPatch<RawMessage> {}\n",
                        encoding="utf-8",
                    )
                    path = str(relative)

                    self.assertEqual(
                        boundaries.audit_keynote_placeholder_visibility_facade_source_topology(
                            root
                        ),
                        sorted(
                            [
                                "focused litchi-keynote placeholder visibility public "
                                f"API exposes raw source bytes source_bytes: {path}:1",
                                "focused litchi-keynote placeholder visibility public "
                                f"API exposes raw byte slice &[u8]: {path}:1",
                                "focused litchi-keynote placeholder visibility public "
                                f"API exposes raw identifier object_id: {path}:1",
                                "focused litchi-keynote placeholder visibility public "
                                f"API exposes archive/IWA type ArchiveObject: {path}:1",
                                "focused litchi-keynote placeholder visibility public "
                                f"API retains flat alias PlaceholderEdit: {path}:2",
                                "focused litchi-keynote placeholder visibility public "
                                f"API exposes wire type WireView: {path}:2",
                                "focused litchi-keynote placeholder visibility public "
                                f"API exposes protobuf type prost: {path}:3",
                                "focused litchi-keynote placeholder visibility public "
                                f"API exposes protobuf type Message: {path}:3",
                                "focused litchi-keynote placeholder visibility public "
                                f"API exposes archive/IWA type RawMessage: {path}:3",
                            ]
                        ),
                    )

    def test_focused_keynote_placeholder_visibility_recursively_scans_future_helpers(
        self,
    ) -> None:
        for helper_root in (
            boundaries.KEYNOTE_PLACEHOLDER_VISIBILITY_OWNER_HELPER_ROOT,
            boundaries.KEYNOTE_PLACEHOLDER_VISIBILITY_PREVIEW_HELPER_ROOT,
        ):
            with self.subTest(helper_root=helper_root):
                with tempfile.TemporaryDirectory() as directory:
                    root = Path(directory)
                    add_keynote_placeholder_visibility_canonical_scaffold(root)
                    nested_relative = helper_root / "future" / "nested.rs"
                    nested = root / nested_relative
                    nested.parent.mkdir(parents=True, exist_ok=True)
                    nested.write_text(
                        "pub type FutureVisibilitySnapshot = ArchiveObject;\n",
                        encoding="utf-8",
                    )

                    self.assertEqual(
                        boundaries.audit_keynote_placeholder_visibility_facade_source_topology(
                            root
                        ),
                        [
                            "focused litchi-keynote placeholder visibility public "
                            "API exposes archive/IWA type ArchiveObject: "
                            f"{nested_relative}:1"
                        ],
                    )

    def test_focused_keynote_placeholder_visibility_allows_private_helper_vocabulary(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            add_keynote_placeholder_visibility_canonical_scaffold(root)
            helpers = (
                boundaries.KEYNOTE_PLACEHOLDER_VISIBILITY_OWNER_HELPER_SOURCES
                + boundaries.KEYNOTE_PLACEHOLDER_VISIBILITY_PREVIEW_HELPER_SOURCES
            )
            for index, relative in enumerate(helpers):
                helper = root / relative
                helper.parent.mkdir(parents=True, exist_ok=True)
                helper.write_text(
                    "fn source_bytes(source_bytes: &[u8], object_id: u64) "
                    "-> ArchiveObject { todo!() }\n"
                    "pub(crate) fn restricted(value: WireView) {}\n"
                    "pub(super) type PlaceholderEdit = RawMessage;\n"
                    "pub(in crate) struct SlideTextPlaceholderPatch;\n"
                    "impl PlaceholderPatch {}\n"
                    "// pub type PlaceholderCommit = ArchiveObject;\n"
                    'const NOTE: &str = "pub fn source_bytes(value: &[u8]) {}";\n'
                    "/* pub type PlaceholderError = WireError; */\n"
                    f"struct Helper{index};\n",
                    encoding="utf-8",
                )
            for helper_root in (
                boundaries.KEYNOTE_PLACEHOLDER_VISIBILITY_OWNER_HELPER_ROOT,
                boundaries.KEYNOTE_PLACEHOLDER_VISIBILITY_PREVIEW_HELPER_ROOT,
            ):
                non_rust = root / helper_root / "future.txt"
                non_rust.write_text(
                    "pub type PlaceholderEdit = ArchiveObject;\n"
                    "pub fn source_bytes(value: &[u8]) {}\n",
                    encoding="utf-8",
                )

            self.assertEqual(
                boundaries.audit_keynote_placeholder_visibility_facade_source_topology(
                    root
                ),
                [],
            )

    def test_focused_keynote_placeholder_visibility_allows_nested_semantic_surface(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            semantic = (
                root / boundaries.KEYNOTE_PLACEHOLDER_VISIBILITY_SEMANTIC_SOURCE
            )
            semantic.parent.mkdir(parents=True)
            semantic.write_text(
                "// pub type PlaceholderEdit = DocumentArchive;\n"
                'const NOTE: &str = "pub struct SlidePlaceholderPatch;";\n'
                "/* pub type SlideTextPlaceholderCommit = buffa::ArchiveView; */\n"
                "pub struct AggregateVisibility;\n"
                "pub struct LayoutVisibility;\n"
                "pub struct SlideNumberVisibilityPolicy;\n"
                "pub fn title_visible(state: &State) -> bool { todo!() }\n"
                "pub fn body_visible(state: &State) -> bool { todo!() }\n"
                "mod slide_number;\n"
                "pub(crate) mod r#slide_number;\n"
                "// pub mod slide_number;\n"
                'const MODULE_NOTE: &str = "pub mod r#slide_number {};";\n'
                "fn source_bytes(source_bytes: &[u8], object_id: u64) "
                "-> DocumentArchive { todo!() }\n"
                "pub(crate) fn restricted(value: ArchiveObject) {}\n"
                "impl PlaceholderEdit {}\n",
                encoding="utf-8",
            )
            owner = root / boundaries.KEYNOTE_PLACEHOLDER_VISIBILITY_OWNER_SOURCE
            owner.parent.mkdir(parents=True, exist_ok=True)
            private_aliases = [
                ("struct" if index % 2 == 0 else "pub(crate) struct")
                + f" {name};"
                for index, name in enumerate(
                    sorted(boundaries.KEYNOTE_PLACEHOLDER_VISIBILITY_FLAT_ALIASES)
                )
            ]
            private_physical = [
                f"pub(super) type Private{index} = {name};"
                for index, name in enumerate(
                    sorted(
                        boundaries.KEYNOTE_PLACEHOLDER_VISIBILITY_PHYSICAL_TYPES
                        | boundaries.KEYNOTE_PLACEHOLDER_VISIBILITY_WIRE_TYPES
                    )
                )
            ]
            private_number_surface = [
                *[
                    ("struct" if index % 2 == 0 else "pub(crate) struct")
                    + f" {name};"
                    for index, name in enumerate(
                        sorted(
                            boundaries.KEYNOTE_PLACEHOLDER_VISIBILITY_FLAT_SEMANTIC_ALIASES
                        )
                    )
                ],
                *[
                    ("fn" if index % 2 == 0 else "pub(super) fn")
                    + f" {name}() {{}}"
                    for index, name in enumerate(
                        sorted(
                            boundaries.KEYNOTE_PLACEHOLDER_VISIBILITY_SLIDE_NUMBER_PUBLIC_MEMBERS
                        )
                    )
                ],
            ]
            owner.write_text(
                "\n".join(
                    [
                        "pub fn slide_placeholder_visibility() "
                        "-> slide::placeholder::State { todo!() }",
                        "pub fn edit_slide_placeholder_visibility("
                        "edit: slide::placeholder::Edit) "
                        "-> slide::placeholder::Diagnostics { todo!() }",
                        "pub fn apply_slide_placeholder_visibility("
                        "patch: slide::placeholder::Patch) "
                        "-> Result<slide::placeholder::Commit, "
                        "slide::placeholder::Error> { todo!() }",
                        *private_aliases,
                        *private_physical,
                        *private_number_surface,
                        "mod slide_number;",
                        "pub(crate) mod r#slide_number;",
                        "// pub mod slide_number;",
                        'const MODULE_NOTE: &str = "pub mod slide_number {};";',
                        "impl SlidePlaceholderVisibilityEdit {}",
                    ]
                )
                + "\n",
                encoding="utf-8",
            )
            lib_export, package_export, slide_export = (
                root / path
                for path in boundaries.KEYNOTE_PLACEHOLDER_VISIBILITY_EXPORT_SOURCES
            )
            private_roots = "\n".join(
                ("use" if index % 2 == 0 else "pub(crate) use")
                + f" crate::slide::placeholder::{name};"
                for index, name in enumerate(
                    sorted(boundaries.KEYNOTE_PLACEHOLDER_VISIBILITY_SHORT_NAMES)
                )
            )
            lib_export.write_text(
                "// pub use crate::slide::placeholder::{Kind, State};\n"
                'const GLOB_NOTE: &str = "pub use crate::slide::placeholder::*;";\n'
                + private_roots
                + "\npub(crate) use crate::slide::placeholder::*;\n"
                "use crate::slide_number as private_number_owner;\n"
                "pub(crate) use crate::r#slide_number as restricted_number_owner;\n"
                "pub(crate) mod slide_number;\n"
                "pub use crate::layout::{Kind, State, Edit, Patch, Commit, "
                "Diagnostics, Error, LimitKind};\n"
                "pub use crate::slide::{SlideInfo, SlideLayout, "
                "SlideNumberVisibilityPolicy};\n"
                "pub fn unrelated(object_id: u64) -> DocumentArchive { todo!() }\n",
                encoding="utf-8",
            )
            package_export.write_text(
                "mod slide_placeholder_visibility;\n"
                "pub(crate) mod r#slide_placeholder_visibility;\n"
                "pub(super) mod slide_number;\n"
                "pub(in crate) use crate::slide_number as private_number_owner;\n"
                "pub(super) use crate::slide::placeholder::{Kind, State};\n"
                "pub fn package_layout(object_id: u64) -> DocumentArchive { todo!() }\n",
                encoding="utf-8",
            )
            slide_export.write_text(
                "pub mod placeholder;\n"
                "mod slide_number;\n"
                "pub(crate) use crate::slide_number as private_number_owner;\n"
                "pub use crate::layout::{Kind, State};\n"
                "pub struct SlideAggregate;\n"
                "pub struct SlideNumberSettings;\n",
                encoding="utf-8",
            )
            preview_parent = (
                root / boundaries.KEYNOTE_PLACEHOLDER_VISIBILITY_SLIDE_PREVIEW_SOURCE
            )
            preview_parent.write_text(
                "mod slide_number;\n"
                "pub(crate) mod r#slide_number;\n"
                "// pub mod slide_number;\n"
                'const NOTE: &str = "pub mod slide_number {};";\n',
                encoding="utf-8",
            )
            nonfocused = root / boundaries.KEYNOTE_SOURCE_ROOT / "render.rs"
            nonfocused.write_text(
                "\n".join(
                    f"pub struct {name};"
                    for name in sorted(
                        boundaries.KEYNOTE_PLACEHOLDER_VISIBILITY_FLAT_ALIASES
                    )
                )
                + "\npub struct SlideNumberVisibility;\n"
                "pub fn show_slide_number() {}\n"
                "pub mod slide_number;\n"
                "pub fn render(object_id: u64) -> DocumentArchive { todo!() }\n",
                encoding="utf-8",
            )
            other_owner = root / "crates/litchi-pages/src/placeholder.rs"
            other_owner.parent.mkdir(parents=True)
            other_owner.write_text(
                "\n".join(
                    f"pub struct {name};"
                    for name in sorted(
                        boundaries.KEYNOTE_PLACEHOLDER_VISIBILITY_FLAT_ALIASES
                    )
                )
                + "\npub use crate::slide::placeholder::*;\n"
                "pub struct SlideNumberVisibility;\n"
                "pub fn hide_slide_number() {}\n"
                "pub mod slide_number;\n"
                "pub fn placeholder(object_id: u64) -> DocumentArchive { todo!() }\n",
                encoding="utf-8",
            )
            add_keynote_placeholder_visibility_canonical_scaffold(root)

            self.assertEqual(
                boundaries.audit_keynote_placeholder_visibility_facade_source_topology(
                    root
                ),
                [],
            )

    def test_numbers_sheet_order_boundary_inventories_are_exact(self) -> None:
        self.assertEqual(
            boundaries.RETIRED_IWA_NUMBERS_SHEET_ORDER_METHODS,
            ("move_sheet",),
        )
        self.assertEqual(
            boundaries.RETIRED_IWA_NUMBERS_SHEET_ORDER_EXAMPLE,
            Path("crates/litchi-iwa/examples/move_numbers_sheet.rs"),
        )
        self.assertEqual(
            boundaries.RETIRED_IWA_NUMBERS_SHEET_ORDER_TESTS,
            (
                "reorders_and_removes_sheets_transactionally",
                "sheet_list_crud_preserves_raw_references_and_restores_exact_component",
                "duplicate_sheet_references_fail_transactionally",
            ),
        )
        self.assertEqual(
            boundaries.IWA_NUMBERS_EDITOR_TEST_SOURCE,
            Path("crates/litchi-iwa/src/numbers/editor/tests.rs"),
        )
        self.assertEqual(
            boundaries.NUMBERS_SHEET_ORDER_SEMANTIC_SOURCE,
            Path("crates/litchi-numbers/src/sheet/order.rs"),
        )
        self.assertEqual(
            boundaries.NUMBERS_SHEET_ORDER_OWNER_SOURCE,
            Path("crates/litchi-numbers/src/package/sheet_order.rs"),
        )
        self.assertEqual(
            boundaries.NUMBERS_SHEET_ORDER_OWNER_HELPER_ROOT,
            Path("crates/litchi-numbers/src/package/sheet_order"),
        )
        self.assertEqual(
            boundaries.NUMBERS_SHEET_ORDER_OWNER_HELPER_SOURCES,
            (
                Path("crates/litchi-numbers/src/package/sheet_order/error.rs"),
                Path("crates/litchi-numbers/src/package/sheet_order/resolve.rs"),
                Path("crates/litchi-numbers/src/package/sheet_order/rewrite.rs"),
            ),
        )
        self.assertEqual(
            boundaries.NUMBERS_SHEET_ORDER_IMPLEMENTATION_SOURCES,
            (
                Path("crates/litchi-numbers/src/sheet/order.rs"),
                Path("crates/litchi-numbers/src/package/sheet_order.rs"),
                *boundaries.NUMBERS_SHEET_ORDER_OWNER_HELPER_SOURCES,
            ),
        )
        self.assertEqual(
            boundaries.NUMBERS_SHEET_ORDER_EXPORT_SOURCES,
            (
                Path("crates/litchi-numbers/src/lib.rs"),
                Path("crates/litchi-numbers/src/package.rs"),
                Path("crates/litchi-numbers/src/sheet.rs"),
            ),
        )
        self.assertEqual(
            boundaries.NUMBERS_SHEET_ORDER_CANONICAL_TYPES,
            ("Edit", "Patch", "Commit", "Diagnostics", "Error", "LimitKind"),
        )
        self.assertEqual(
            boundaries.NUMBERS_SHEET_ORDER_PACKAGE_METHODS,
            ("edit_sheet_order", "apply_sheet_order"),
        )
        self.assertEqual(
            boundaries.NUMBERS_SHEET_ORDER_FLAT_ALIASES,
            frozenset(
                {
                    "SheetOrderEdit",
                    "SheetOrderPatch",
                    "SheetOrderCommit",
                    "SheetOrderDiagnostics",
                    "SheetOrderError",
                    "SheetOrderLimitKind",
                }
            ),
        )
        self.assertEqual(
            boundaries.NUMBERS_SHEET_ORDER_PHYSICAL_TYPES,
            frozenset(
                {
                    "Archive",
                    "ArchiveObject",
                    "ComponentCatalog",
                    "DocumentSnapshot",
                    "EntryEdit",
                    "ExactArtifacts",
                    "IWorkPackage",
                    "PhysicalSource",
                    "RawMessage",
                    "Resolved",
                    "SheetOrderSnapshot",
                    "SheetSnapshot",
                    "SnappyStream",
                    "SourceCatalog",
                }
            ),
        )
        self.assertEqual(
            boundaries.NUMBERS_SHEET_ORDER_WIRE_TYPES,
            frozenset(
                {
                    "DecodeOptions",
                    "NestedFieldEdit",
                    "NestedFieldReplacement",
                    "WireDescent",
                    "WireError",
                    "WireLimits",
                    "WireResourceLimit",
                    "WireView",
                }
            ),
        )
        self.assertEqual(
            boundaries.NUMBERS_SHEET_ORDER_PROTO_ORIGINS,
            frozenset({"tn", "tsp"}),
        )

    def test_retired_iwa_numbers_sheet_order_surface_cannot_return(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            workbook = root / boundaries.IWA_NUMBERS_SEMANTIC_WORKBOOK_SOURCE
            workbook.parent.mkdir(parents=True)
            workbook.write_text("pub fn r#move_sheet() {}\n", encoding="utf-8")
            tests = root / boundaries.IWA_NUMBERS_EDITOR_TEST_SOURCE
            tests.parent.mkdir(parents=True, exist_ok=True)
            tests.write_text(
                "fn reorders_and_removes_sheets_transactionally() {}\n"
                "fn sheet_list_crud_preserves_raw_references_and_restores_exact_component() {}\n"
                "fn duplicate_sheet_references_fail_transactionally() {}\n",
                encoding="utf-8",
            )
            example = root / boundaries.RETIRED_IWA_NUMBERS_SHEET_ORDER_EXAMPLE
            example.parent.mkdir(parents=True)
            example.write_text("fn main() {}\n", encoding="utf-8")

            self.assertEqual(
                boundaries.audit_iwa_numbers_sheet_order_source_topology(root),
                sorted(
                    [
                        "retired litchi-iwa Numbers sheet-order example returned: "
                        "crates/litchi-iwa/examples/move_numbers_sheet.rs",
                        "retired litchi-iwa Numbers sheet-order method move_sheet: "
                        "crates/litchi-iwa/src/numbers/editor/semantic/workbook.rs:1",
                        "retired litchi-iwa Numbers sheet-order test "
                        "reorders_and_removes_sheets_transactionally: "
                        "crates/litchi-iwa/src/numbers/editor/tests.rs:1",
                        "retired litchi-iwa Numbers sheet-order test "
                        "sheet_list_crud_preserves_raw_references_and_restores_exact_"
                        "component: crates/litchi-iwa/src/numbers/editor/tests.rs:2",
                        "retired litchi-iwa Numbers sheet-order test "
                        "duplicate_sheet_references_fail_transactionally: "
                        "crates/litchi-iwa/src/numbers/editor/tests.rs:3",
                    ]
                ),
            )

    def test_retired_iwa_numbers_sheet_order_readme_calls_and_example(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            readme = root / boundaries.IWA_NUMBERS_README
            readme.parent.mkdir(parents=True)
            readme.write_text(
                "numbers.r#move_sheet(0, 1);\n"
                "numbers_editor\n  .\n  move_sheet(1, 0);\n"
                "crate::nested::NumbersEditor::\n  r#move_sheet(0, 1);\n"
                "move_numbers_sheet\n"
                "move_numbers_sheet.rs\n",
                encoding="utf-8",
            )
            self.assertEqual(
                boundaries.audit_iwa_numbers_sheet_order_source_topology(root),
                sorted(
                    [
                        "retired litchi-iwa Numbers sheet-order README call move_sheet: "
                        "crates/litchi-iwa/README.md:1",
                        "retired litchi-iwa Numbers sheet-order README call move_sheet: "
                        "crates/litchi-iwa/README.md:4",
                        "retired litchi-iwa Numbers sheet-order README call move_sheet: "
                        "crates/litchi-iwa/README.md:6",
                        "retired litchi-iwa Numbers sheet-order README example reference "
                        "move_numbers_sheet: crates/litchi-iwa/README.md:7",
                        "retired litchi-iwa Numbers sheet-order README example reference "
                        "move_numbers_sheet: crates/litchi-iwa/README.md:8",
                    ]
                ),
            )

    def test_iwa_numbers_sheet_order_policy_retains_other_sheet_operations(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            workbook = root / boundaries.IWA_NUMBERS_SEMANTIC_WORKBOOK_SOURCE
            workbook.parent.mkdir(parents=True)
            workbook.write_text(
                "// fn move_sheet() {}\n"
                'const NOTE: &str = "fn r#move_sheet() {}";\n'
                "/* fn move_sheet() {} */\n"
                "pub fn move_sheet_drawable() {}\n"
                "pub fn move_table() {}\n"
                "pub fn add_empty_sheet() {}\n"
                "pub fn duplicate_sheet() {}\n"
                "pub fn remove_sheet() {}\n"
                "pub fn update_numbers_document() {}\n",
                encoding="utf-8",
            )
            tests = root / boundaries.IWA_NUMBERS_EDITOR_TEST_SOURCE
            tests.parent.mkdir(parents=True, exist_ok=True)
            tests.write_text(
                "// fn reorders_and_removes_sheets_transactionally() {}\n"
                'const NOTE: &str = "fn duplicate_sheet_references_fail_transactionally() {}";\n'
                "fn reorders_and_removes_sheets_transactionally_v2() {}\n",
                encoding="utf-8",
            )
            other_owner = root / "crates/litchi-pages/src/editor.rs"
            other_owner.parent.mkdir(parents=True)
            other_owner.write_text("pub fn move_sheet() {}\n", encoding="utf-8")
            retained = root / "crates/litchi-iwa/examples/create_numbers_sheets.rs"
            retained.parent.mkdir(parents=True)
            retained.write_text("fn main() {}\n", encoding="utf-8")
            readme = root / boundaries.IWA_NUMBERS_README
            readme.parent.mkdir(parents=True, exist_ok=True)
            readme.write_text(
                "`move_sheet` is retired prose.\n"
                "move_sheet\n"
                "edit.move_sheet(selector, 1)\n"
                "package.edit_sheet_order()\n"
                "package.apply_sheet_order(patch)\n"
                "move_numbers_sheets.rs\n",
                encoding="utf-8",
            )
            other_readme = root / "README.md"
            other_readme.write_text("numbers.move_sheet(0, 1)\n", encoding="utf-8")

            self.assertEqual(
                boundaries.audit_iwa_numbers_sheet_order_source_topology(root),
                [],
            )

    def test_focused_numbers_sheet_order_requires_each_canonical_type(self) -> None:
        for missing in boundaries.NUMBERS_SHEET_ORDER_CANONICAL_TYPES:
            with self.subTest(missing=missing):
                with tempfile.TemporaryDirectory() as directory:
                    root = Path(directory)
                    add_numbers_sheet_order_canonical_scaffold(root)
                    semantic = root / boundaries.NUMBERS_SHEET_ORDER_SEMANTIC_SOURCE
                    semantic.write_text(
                        "".join(
                            f"pub struct {name};\n"
                            for name in boundaries.NUMBERS_SHEET_ORDER_CANONICAL_TYPES
                            if name != missing
                        ),
                        encoding="utf-8",
                    )
                    self.assertEqual(
                        boundaries.audit_numbers_sheet_order_facade_source_topology(root),
                        [
                            "focused litchi-numbers sheet-order public API is missing "
                            f"canonical sheet::order type {missing}: "
                            "crates/litchi-numbers/src/sheet/order.rs"
                        ],
                    )

    def test_focused_numbers_sheet_order_requires_nested_modules_and_private_owner(
        self,
    ) -> None:
        missing_root = (
            "",
            "mod sheet;\n",
            "pub(crate) mod sheet;\n",
            "pub(super) mod r#sheet {}\n",
            "pub(in crate) mod sheet;\n",
            "// pub mod sheet;\n",
            'const NOTE: &str = "pub mod sheet;";\n',
        )
        accepted_root = (
            "pub mod sheet;\n",
            "pub mod r#sheet;\n",
            "pub mod sheet {}\n",
            "pub\nmod\nr#sheet\n{}\n",
        )
        for declaration, expected in (
            *[
                (
                    declaration,
                    [
                        "focused litchi-numbers sheet-order public API is missing "
                        "canonical root sheet module: crates/litchi-numbers/src/lib.rs"
                    ],
                )
                for declaration in missing_root
            ],
            *[(declaration, []) for declaration in accepted_root],
        ):
            with self.subTest(scope="root", declaration=declaration):
                with tempfile.TemporaryDirectory() as directory:
                    root = Path(directory)
                    add_numbers_sheet_order_canonical_scaffold(root)
                    path = root / boundaries.NUMBERS_SHEET_ORDER_EXPORT_SOURCES[0]
                    path.write_text(declaration, encoding="utf-8")
                    self.assertEqual(
                        boundaries.audit_numbers_sheet_order_facade_source_topology(root),
                        expected,
                    )

        missing_nested = (
            "",
            "mod order;\n",
            "pub(crate) mod order;\n",
            "pub(super) mod r#order {}\n",
            "pub(in crate) mod order;\n",
            "// pub mod order;\n",
            'const NOTE: &str = "pub mod order;";\n',
        )
        accepted_nested = (
            "pub mod order;\n",
            "pub mod r#order;\n",
            "pub mod order {}\n",
            "pub\nmod\nr#order\n{}\n",
        )
        for declaration, expected in (
            *[
                (
                    declaration,
                    [
                        "focused litchi-numbers sheet-order public API is missing "
                        "canonical sheet::order module: crates/litchi-numbers/src/sheet.rs"
                    ],
                )
                for declaration in missing_nested
            ],
            *[(declaration, []) for declaration in accepted_nested],
        ):
            with self.subTest(scope="nested", declaration=declaration):
                with tempfile.TemporaryDirectory() as directory:
                    root = Path(directory)
                    add_numbers_sheet_order_canonical_scaffold(root)
                    path = root / boundaries.NUMBERS_SHEET_ORDER_EXPORT_SOURCES[2]
                    path.write_text(declaration, encoding="utf-8")
                    self.assertEqual(
                        boundaries.audit_numbers_sheet_order_facade_source_topology(root),
                        expected,
                    )

        missing_owner = (
            "",
            "// mod sheet_order;\n",
            'const NOTE: &str = "mod sheet_order;";\n',
            "mod sheet_orders;\n",
        )
        accepted_owner = (
            "mod sheet_order;\n",
            "mod r#sheet_order {}\n",
            "pub(crate) mod sheet_order;\n",
            "pub(super) mod r#sheet_order {}\n",
            "pub(in crate)\nmod\nsheet_order\n;\n",
        )
        for declaration, expected in (
            *[
                (
                    declaration,
                    [
                        "focused litchi-numbers sheet-order public API is missing "
                        "private package owner module: "
                        "crates/litchi-numbers/src/package.rs"
                    ],
                )
                for declaration in missing_owner
            ],
            *[(declaration, []) for declaration in accepted_owner],
        ):
            with self.subTest(scope="owner", declaration=declaration):
                with tempfile.TemporaryDirectory() as directory:
                    root = Path(directory)
                    add_numbers_sheet_order_canonical_scaffold(root)
                    path = root / boundaries.NUMBERS_SHEET_ORDER_EXPORT_SOURCES[1]
                    path.write_text(declaration, encoding="utf-8")
                    self.assertEqual(
                        boundaries.audit_numbers_sheet_order_facade_source_topology(root),
                        expected,
                    )

        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            add_numbers_sheet_order_canonical_scaffold(root)
            (root / boundaries.NUMBERS_SHEET_ORDER_OWNER_SOURCE).unlink()
            self.assertEqual(
                boundaries.audit_numbers_sheet_order_facade_source_topology(root),
                [
                    "focused litchi-numbers sheet-order public API is missing "
                    "private package owner source: "
                    "crates/litchi-numbers/src/package/sheet_order.rs"
                ],
            )

    def test_focused_numbers_sheet_order_rejects_duplicate_package_module(self) -> None:
        for declaration in (
            "pub mod sheet_order;",
            "pub mod r#sheet_order;",
            "pub mod sheet_order {}",
            "pub\nmod\nr#sheet_order\n{}",
        ):
            with self.subTest(declaration=declaration):
                with tempfile.TemporaryDirectory() as directory:
                    root = Path(directory)
                    add_numbers_sheet_order_canonical_scaffold(root)
                    package = root / boundaries.NUMBERS_SHEET_ORDER_EXPORT_SOURCES[1]
                    package.write_text(declaration + "\n", encoding="utf-8")
                    self.assertEqual(
                        boundaries.audit_numbers_sheet_order_facade_source_topology(root),
                        [
                            "focused litchi-numbers sheet-order public API exposes "
                            "duplicate package::sheet_order module: "
                            "crates/litchi-numbers/src/package.rs:1"
                        ],
                    )

    def test_focused_numbers_sheet_order_rejects_all_flat_aliases(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            relative_sources = (
                boundaries.NUMBERS_SHEET_ORDER_IMPLEMENTATION_SOURCES
                + boundaries.NUMBERS_SHEET_ORDER_EXPORT_SOURCES
            )
            sources = tuple(root / path for path in relative_sources)
            declarations: dict[Path, list[str]] = {path: [] for path in sources}
            expected: list[str] = []
            for index, name in enumerate(sorted(boundaries.NUMBERS_SHEET_ORDER_FLAT_ALIASES)):
                source_index = index % len(sources)
                path = sources[source_index]
                declaration = (
                    f"pub struct {name};"
                    if source_index < 2
                    else f"pub use crate::legacy::Legacy as {name};"
                )
                declarations[path].append(declaration)
                expected.append(
                    "focused litchi-numbers sheet-order public API retains flat alias "
                    f"{name}: {relative_sources[source_index]}:"
                    f"{len(declarations[path])}"
                )
            for path, lines in declarations.items():
                path.parent.mkdir(parents=True, exist_ok=True)
                path.write_text("\n".join(lines) + "\n", encoding="utf-8")
            add_numbers_sheet_order_canonical_scaffold(root)
            self.assertEqual(
                boundaries.audit_numbers_sheet_order_facade_source_topology(root),
                sorted(expected),
            )

    def test_focused_numbers_sheet_order_rejects_root_aliases_glob_and_owner_aliases(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            lib, package, sheet = (
                root / path for path in boundaries.NUMBERS_SHEET_ORDER_EXPORT_SOURCES
            )
            lib.parent.mkdir(parents=True)
            aliases = sorted(boundaries.NUMBERS_SHEET_ORDER_SHORT_NAMES)
            lib_names, package_names, sheet_names = aliases[:2], aliases[2:4], aliases[4:]
            lib.write_text(
                "pub use crate::sheet::order::{" + ", ".join(lib_names) + "};\n",
                encoding="utf-8",
            )
            package.write_text(
                "pub use crate::sheet_order::{" + ", ".join(package_names) + "};\n"
                "pub use crate::sheet::order::*;\n",
                encoding="utf-8",
            )
            sheet.write_text(
                "pub use crate::sheet::order::{" + ", ".join(sheet_names) + "};\n",
                encoding="utf-8",
            )
            add_numbers_sheet_order_canonical_scaffold(root)
            expected = [
                *[
                    "focused litchi-numbers sheet-order public API retains root alias "
                    f"{name}: crates/litchi-numbers/src/lib.rs:1"
                    for name in lib_names
                ],
                *[
                    "focused litchi-numbers sheet-order public API retains root alias "
                    f"{name}: crates/litchi-numbers/src/package.rs:1"
                    for name in package_names
                ],
                *[
                    "focused litchi-numbers sheet-order public API retains root alias "
                    f"{name}: crates/litchi-numbers/src/sheet.rs:1"
                    for name in sheet_names
                ],
                "focused litchi-numbers sheet-order public API exposes public "
                "sheet-order owner alias: crates/litchi-numbers/src/lib.rs:1",
                "focused litchi-numbers sheet-order public API exposes public "
                "sheet-order owner alias: crates/litchi-numbers/src/package.rs:1",
                "focused litchi-numbers sheet-order public API exposes public "
                "sheet-order owner alias: crates/litchi-numbers/src/package.rs:2",
                "focused litchi-numbers sheet-order public API exposes public "
                "sheet-order owner alias: crates/litchi-numbers/src/sheet.rs:1",
                "focused litchi-numbers sheet-order public API retains root aliases "
                "via sheet::order glob: crates/litchi-numbers/src/package.rs:2",
            ]
            self.assertEqual(
                boundaries.audit_numbers_sheet_order_facade_source_topology(root),
                sorted(expected),
            )

    def test_focused_numbers_sheet_order_rejects_public_owner_alias_variants(
        self,
    ) -> None:
        for relative in boundaries.NUMBERS_SHEET_ORDER_EXPORT_SOURCES:
            for declaration in (
                "pub use crate::sheet::order as ordering;",
                "pub use crate::r#sheet_order as ordering;",
                "pub\nuse\ncrate::sheet::order as ordering;",
                "pub type sheet_order = bool;",
            ):
                with self.subTest(relative=relative, declaration=declaration):
                    with tempfile.TemporaryDirectory() as directory:
                        root = Path(directory)
                        add_numbers_sheet_order_canonical_scaffold(root)
                        path = root / relative
                        source = path.read_text(encoding="utf-8")
                        line = source.count("\n") + 1
                        path.write_text(source + declaration + "\n", encoding="utf-8")
                        self.assertEqual(
                            boundaries.audit_numbers_sheet_order_facade_source_topology(
                                root
                            ),
                            [
                                "focused litchi-numbers sheet-order public API exposes "
                                f"public sheet-order owner alias: {relative}:{line}"
                            ],
                        )

    def test_focused_numbers_sheet_order_rejects_physical_leaks(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            semantic = root / boundaries.NUMBERS_SHEET_ORDER_SEMANTIC_SOURCE
            semantic.parent.mkdir(parents=True)
            semantic.write_text(
                "pub fn order(r#source_bytes: &[u8], r#object_id: u64) "
                "-> DocumentArchive { todo!() }\n"
                "pub type Projection = buffa::DocumentArchiveView;\n"
                "impl prost::Message for order::Edit {}\n",
                encoding="utf-8",
            )
            owner = root / boundaries.NUMBERS_SHEET_ORDER_OWNER_SOURCE
            owner.parent.mkdir(parents=True, exist_ok=True)
            owner_lines = [
                f"pub type Physical{index} = {name};"
                for index, name in enumerate(
                    sorted(boundaries.NUMBERS_SHEET_ORDER_PHYSICAL_TYPES)
                )
            ]
            owner_lines.extend(
                f"pub type WireLeak{index} = {name};"
                for index, name in enumerate(
                    sorted(boundaries.NUMBERS_SHEET_ORDER_WIRE_TYPES)
                )
            )
            owner_lines.extend(
                [
                    "pub type ProtoProjection = (tn::DocumentArchive, "
                    "tsp::ReferenceArchive);",
                    "pub fn physical_names(component_name: &str, "
                    "member_name: &str, entry_name: &str) {}",
                ]
            )
            owner.write_text("\n".join(owner_lines) + "\n", encoding="utf-8")
            lib, package, sheet = (
                root / path for path in boundaries.NUMBERS_SHEET_ORDER_EXPORT_SOURCES
            )
            lib.write_text(
                "pub fn edit_sheet_order(value: litchi_iwa_protos::GeneratedSheetOrder) "
                "-> order::Edit { todo!() }\n",
                encoding="utf-8",
            )
            package.write_text(
                "pub fn apply_sheet_order(value: prost_types::MessageInfo) "
                "-> order::Commit { todo!() }\n",
                encoding="utf-8",
            )
            sheet.write_text(
                "pub fn edit_sheet_order() -> SourceBytes { todo!() }\n",
                encoding="utf-8",
            )
            add_numbers_sheet_order_canonical_scaffold(root)

            violations = boundaries.audit_numbers_sheet_order_facade_source_topology(root)
            self.assertEqual(violations, sorted(violations))
            self.assertTrue(
                all(
                    violation.startswith(
                        "focused litchi-numbers sheet-order public API exposes "
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
                "protobuf type tn",
                "protobuf type tsp",
                "protobuf type ProtoProjection",
                "physical package name component_name",
                "physical package name member_name",
                "physical package name entry_name",
                "archive/IWA type litchi_iwa_protos",
                "generated type GeneratedSheetOrder",
                "protobuf type prost_types",
                "archive/IWA type MessageInfo",
                "raw source bytes SourceBytes",
                *tuple(
                    f"archive/IWA type {name}"
                    for name in sorted(boundaries.NUMBERS_SHEET_ORDER_PHYSICAL_TYPES)
                ),
                *tuple(
                    f"wire type {name}"
                    for name in sorted(boundaries.NUMBERS_SHEET_ORDER_WIRE_TYPES)
                ),
            )
            self.assertEqual(len(violations), 43)
            for fragment in expected_fragments:
                with self.subTest(fragment=fragment):
                    self.assertTrue(
                        any(fragment in violation for violation in violations),
                        msg=f"missing focused sheet-order leak: {fragment}",
                    )

    def test_focused_numbers_sheet_order_recursively_scans_private_helpers(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            add_numbers_sheet_order_canonical_scaffold(root)
            nested_relative = (
                boundaries.NUMBERS_SHEET_ORDER_OWNER_HELPER_ROOT
                / "future"
                / "nested.rs"
            )
            nested = root / nested_relative
            nested.parent.mkdir(parents=True, exist_ok=True)
            nested.write_text(
                "pub fn expose(source_bytes: &[u8], object_id: u64) "
                "-> ArchiveObject { todo!() }\n"
                "pub type SheetOrderEdit = WireView;\n"
                "impl prost::Message for SheetOrderPatch<RawMessage> {}\n",
                encoding="utf-8",
            )
            path = str(nested_relative)
            self.assertEqual(
                boundaries.audit_numbers_sheet_order_facade_source_topology(root),
                sorted(
                    [
                        "focused litchi-numbers sheet-order public API exposes "
                        f"raw source bytes source_bytes: {path}:1",
                        "focused litchi-numbers sheet-order public API exposes "
                        f"raw byte slice &[u8]: {path}:1",
                        "focused litchi-numbers sheet-order public API exposes "
                        f"raw identifier object_id: {path}:1",
                        "focused litchi-numbers sheet-order public API exposes "
                        f"archive/IWA type ArchiveObject: {path}:1",
                        "focused litchi-numbers sheet-order public API retains flat "
                        f"alias SheetOrderEdit: {path}:2",
                        "focused litchi-numbers sheet-order public API exposes "
                        f"wire type WireView: {path}:2",
                        "focused litchi-numbers sheet-order public API exposes "
                        f"protobuf type prost: {path}:3",
                        "focused litchi-numbers sheet-order public API exposes "
                        f"protobuf type Message: {path}:3",
                        "focused litchi-numbers sheet-order public API exposes "
                        f"archive/IWA type RawMessage: {path}:3",
                    ]
                ),
            )

    def test_focused_numbers_sheet_order_allows_canonical_and_retained_surfaces(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            semantic = root / boundaries.NUMBERS_SHEET_ORDER_SEMANTIC_SOURCE
            semantic.parent.mkdir(parents=True)
            semantic.write_text(
                "// pub type SheetOrderEdit = DocumentArchive;\n"
                'const NOTE: &str = "pub struct SheetOrderPatch;";\n'
                "pub struct Sheet;\n"
                "pub struct SheetSelector;\n"
                "pub fn move_sheet(selector: SheetSelector, destination: usize) {}\n"
                "fn source_bytes(source_bytes: &[u8], object_id: u64) "
                "-> ArchiveObject { todo!() }\n"
                "pub(crate) fn restricted(value: WireView) {}\n",
                encoding="utf-8",
            )
            owner = root / boundaries.NUMBERS_SHEET_ORDER_OWNER_SOURCE
            owner.parent.mkdir(parents=True, exist_ok=True)
            private_aliases = [
                ("struct" if index % 2 == 0 else "pub(crate) struct")
                + f" {name};"
                for index, name in enumerate(
                    sorted(boundaries.NUMBERS_SHEET_ORDER_FLAT_ALIASES)
                )
            ]
            private_physical = [
                f"pub(super) type Private{index} = {name};"
                for index, name in enumerate(
                    sorted(
                        boundaries.NUMBERS_SHEET_ORDER_PHYSICAL_TYPES
                        | boundaries.NUMBERS_SHEET_ORDER_WIRE_TYPES
                    )
                )
            ]
            owner.write_text(
                "\n".join(
                    [
                        "pub fn edit_sheet_order() -> sheet::order::Edit { todo!() }",
                        "pub fn apply_sheet_order(patch: sheet::order::Patch) "
                        "-> sheet::order::Commit { todo!() }",
                        "pub fn move_sheet(selector: SheetSelector, destination: usize) {}",
                        *private_aliases,
                        *private_physical,
                        "impl SheetOrderEdit {}",
                    ]
                )
                + "\n",
                encoding="utf-8",
            )
            lib, package, sheet = (
                root / path for path in boundaries.NUMBERS_SHEET_ORDER_EXPORT_SOURCES
            )
            private_roots = "\n".join(
                ("use" if index % 2 == 0 else "pub(crate) use")
                + f" crate::sheet::order::{name};"
                for index, name in enumerate(
                    sorted(boundaries.NUMBERS_SHEET_ORDER_SHORT_NAMES)
                )
            )
            lib.write_text(
                "// pub use crate::sheet::order::{Edit, Patch};\n"
                'const NOTE: &str = "pub use crate::sheet::order::*;";\n'
                + private_roots
                + "\npub(crate) use crate::sheet::order::*;\n"
                "pub use crate::table::{Edit, Patch, Commit, Diagnostics, Error, LimitKind};\n"
                "pub struct Sheet;\n"
                "pub struct SheetSelector;\n",
                encoding="utf-8",
            )
            package.write_text(
                "mod sheet_order;\n"
                "pub(crate) mod r#sheet_order;\n"
                "pub(super) use crate::sheet::order::{Edit, Patch};\n"
                "pub fn move_table() {}\n"
                "pub fn add_empty_sheet() {}\n"
                "pub fn duplicate_sheet() {}\n"
                "pub fn remove_sheet() {}\n",
                encoding="utf-8",
            )
            sheet.write_text(
                "pub mod order;\n"
                "pub struct Sheet;\n"
                "pub struct SheetSelector;\n",
                encoding="utf-8",
            )
            helper = root / boundaries.NUMBERS_SHEET_ORDER_OWNER_HELPER_ROOT / "private.rs"
            helper.parent.mkdir(parents=True, exist_ok=True)
            helper.write_text(
                "fn source_bytes(source_bytes: &[u8], object_id: u64) "
                "-> ArchiveObject { todo!() }\n"
                "pub(crate) type SheetOrderPatch = WireView;\n"
                "impl SheetOrderEdit {}\n",
                encoding="utf-8",
            )
            helper.with_suffix(".txt").write_text(
                "pub type SheetOrderEdit = ArchiveObject;\n",
                encoding="utf-8",
            )
            nonfocused = root / boundaries.NUMBERS_SOURCE_ROOT / "names.rs"
            nonfocused.write_text(
                "\n".join(
                    f"pub struct {name};"
                    for name in sorted(boundaries.NUMBERS_SHEET_ORDER_FLAT_ALIASES)
                )
                + "\npub fn move_sheet(object_id: u64) -> ArchiveObject { todo!() }\n",
                encoding="utf-8",
            )
            other_owner = root / "crates/litchi-pages/src/sheet/order.rs"
            other_owner.parent.mkdir(parents=True)
            other_owner.write_text(
                "pub struct SheetOrderEdit;\n"
                "pub fn move_sheet(object_id: u64) -> ArchiveObject { todo!() }\n",
                encoding="utf-8",
            )
            add_numbers_sheet_order_canonical_scaffold(root)

            self.assertEqual(
                boundaries.audit_numbers_sheet_order_facade_source_topology(root),
                [],
            )

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

    def test_retired_iwa_numbers_names_inventories_are_exact(self) -> None:
        self.assertEqual(
            boundaries.RETIRED_IWA_NUMBERS_NAMES_METHODS,
            ("rename_sheet", "rename_table"),
        )
        self.assertEqual(
            boundaries.IWA_NUMBERS_SEMANTIC_WORKBOOK_SOURCE,
            Path("crates/litchi-iwa/src/numbers/editor/semantic/workbook.rs"),
        )
        self.assertEqual(
            boundaries.RETIRED_IWA_NUMBERS_NAMES_EXAMPLE,
            Path("crates/litchi-iwa/examples/rename_numbers_items.rs"),
        )
        self.assertEqual(
            boundaries.IWA_NUMBERS_README,
            Path("crates/litchi-iwa/README.md"),
        )
        self.assertEqual(
            boundaries.NUMBERS_NAMES_CANONICAL_TYPES,
            ("Edit", "Patch", "Commit", "Diagnostics", "Error", "LimitKind"),
        )
        self.assertEqual(
            boundaries.NUMBERS_NAMES_OPTIONAL_TYPES,
            ("Path", "InvalidReason"),
        )
        self.assertEqual(
            boundaries.NUMBERS_NAMES_FLAT_ALIAS_PREFIXES,
            ("Name", "Names", "SheetName", "TableName"),
        )
        self.assertEqual(
            boundaries.NUMBERS_NAMES_SHORT_NAMES,
            frozenset(
                {
                    "Edit",
                    "Patch",
                    "Commit",
                    "Diagnostics",
                    "Error",
                    "LimitKind",
                    "Path",
                    "InvalidReason",
                }
            ),
        )
        self.assertEqual(
            boundaries.NUMBERS_NAMES_FLAT_ALIASES,
            frozenset(
                prefix + suffix
                for prefix in ("Name", "Names", "SheetName", "TableName")
                for suffix in (
                    "Edit",
                    "Patch",
                    "Commit",
                    "Diagnostics",
                    "Error",
                    "LimitKind",
                    "Path",
                    "InvalidReason",
                )
            ),
        )
        self.assertEqual(
            boundaries.NUMBERS_NAMES_PHYSICAL_TYPES,
            frozenset(
                {
                    "Archive",
                    "ComponentCatalog",
                    "EntryEdit",
                    "RawMessage",
                    "Resolved",
                    "SnappyStream",
                }
            ),
        )
        self.assertEqual(
            boundaries.NUMBERS_NAMES_WIRE_TYPES,
            frozenset(
                {
                    "WireDescent",
                    "WireError",
                    "WireLimits",
                    "WireResourceLimit",
                    "WireView",
                }
            ),
        )

    def test_retired_iwa_numbers_names_host_surface_cannot_return(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            nested = root / boundaries.IWA_NUMBERS_SOURCE_ROOT / "legacy/names.rs"
            nested.parent.mkdir(parents=True)
            nested.write_text(
                "pub(in crate::numbers) async unsafe fn rename_table() {}\n",
                encoding="utf-8",
            )
            workbook = root / boundaries.IWA_NUMBERS_SEMANTIC_WORKBOOK_SOURCE
            workbook.parent.mkdir(parents=True, exist_ok=True)
            workbook.write_text("fn r#rename_sheet() {}\n", encoding="utf-8")
            example = root / boundaries.RETIRED_IWA_NUMBERS_NAMES_EXAMPLE
            example.parent.mkdir(parents=True, exist_ok=True)
            example.write_text("// retired example returned\n", encoding="utf-8")

            self.assertEqual(
                boundaries.audit_iwa_numbers_names_source_topology(root),
                sorted(
                    [
                        "retired litchi-iwa Numbers names method rename_sheet: "
                        "crates/litchi-iwa/src/numbers/editor/semantic/workbook.rs:1",
                        "retired litchi-iwa Numbers names method rename_table: "
                        "crates/litchi-iwa/src/numbers/legacy/names.rs:1",
                        "retired litchi-iwa Numbers names example returned: "
                        "crates/litchi-iwa/examples/rename_numbers_items.rs",
                    ]
                ),
            )

    def test_retired_iwa_numbers_names_readme_calls_and_example_references(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            readme = root / boundaries.IWA_NUMBERS_README
            readme.parent.mkdir(parents=True)
            readme.write_text(
                "\n".join(
                    [
                        "let sheet = numbers",
                        "    .",
                        "    r#rename_sheet",
                        "    (",
                        "numbers_editor.r#rename_table(",
                        "crate::numbers::NumbersEditor",
                        "    ::",
                        "    rename_sheet(",
                        "r#NumbersEditor::r#rename_table(",
                        "Run `rename_numbers_items`.",
                        "cargo run --example rename_numbers_items.rs",
                    ]
                )
                + "\n",
                encoding="utf-8",
            )

            self.assertEqual(
                boundaries.audit_iwa_numbers_names_source_topology(root),
                sorted(
                    [
                        "retired litchi-iwa Numbers names README call "
                        "rename_sheet: crates/litchi-iwa/README.md:3",
                        "retired litchi-iwa Numbers names README call "
                        "rename_table: crates/litchi-iwa/README.md:5",
                        "retired litchi-iwa Numbers names README call "
                        "rename_sheet: crates/litchi-iwa/README.md:8",
                        "retired litchi-iwa Numbers names README call "
                        "rename_table: crates/litchi-iwa/README.md:9",
                        "retired litchi-iwa Numbers names README example reference "
                        "rename_numbers_items: crates/litchi-iwa/README.md:10",
                        "retired litchi-iwa Numbers names README example reference "
                        "rename_numbers_items: crates/litchi-iwa/README.md:11",
                    ]
                ),
            )

    def test_iwa_numbers_names_policy_ignores_trivia_near_names_and_other_owners(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            host = root / boundaries.IWA_NUMBERS_SOURCE_ROOT / "legacy/names_old.rs"
            host.parent.mkdir(parents=True)
            host.write_text(
                "\n".join(
                    [
                        "// pub fn rename_sheet() {}",
                        'const NOTE: &str = "fn rename_table() {}";',
                        "/* fn rename_sheet() {}",
                        "   /* fn rename_table() {} */",
                        "   fn rename_sheet() {} */",
                        'const RAW_NOTE: &str = r###"fn rename_table() {}"###;',
                        "pub fn rename_sheet_snapshot() {}",
                        "pub fn rename_tables() {}",
                        "pub fn legacy_rename_sheet() {}",
                    ]
                )
                + "\n",
                encoding="utf-8",
            )
            readme = root / boundaries.IWA_NUMBERS_README
            readme.parent.mkdir(parents=True, exist_ok=True)
            readme.write_text(
                "\n".join(
                    [
                        "The rename_sheet and rename_table names are retired.",
                        "Use `rename_sheet` rather than `.rename_sheet`.",
                        "rename_sheet(",
                        "rename_table(",
                        "numbers.rename_sheet_snapshot(",
                        "numbers_editor.rename_tables(",
                        "Package::edit_names(",
                        "package.edit_names(",
                        "builder.sheet_name(",
                        "builder.table_name(",
                        "rename_numbers_item",
                        "rename_numbers_items_old.rs",
                    ]
                )
                + "\n",
                encoding="utf-8",
            )
            declarations = "pub fn rename_sheet() {}\npub fn rename_table() {}\n"
            for relative in (
                Path("crates/litchi-numbers/src/names.rs"),
                Path("crates/litchi-iwa/src/pages/editor/tables/semantic.rs"),
                Path("crates/litchi-iwa/src/keynote/editor/slide_tables/semantic.rs"),
            ):
                path = root / relative
                path.parent.mkdir(parents=True, exist_ok=True)
                path.write_text(declarations, encoding="utf-8")
            non_rust = root / boundaries.IWA_NUMBERS_SOURCE_ROOT / "names.txt"
            non_rust.write_text(declarations, encoding="utf-8")
            near_example = (
                root
                / boundaries.RETIRED_IWA_NUMBERS_NAMES_EXAMPLE.with_name(
                    "rename_numbers_items_old.rs"
                )
            )
            near_example.parent.mkdir(parents=True, exist_ok=True)
            near_example.write_text("// near example\n", encoding="utf-8")
            for relative in (
                Path("README.md"),
                Path("crates/litchi-numbers/README.md"),
            ):
                other = root / relative
                other.parent.mkdir(parents=True, exist_ok=True)
                other.write_text(
                    "numbers.rename_sheet(\nrename_numbers_items\n",
                    encoding="utf-8",
                )

            self.assertEqual(
                boundaries.audit_iwa_numbers_names_source_topology(root), []
            )

    def test_focused_numbers_names_public_api_rejects_physical_leaks(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            semantic, transaction = (
                root / path for path in boundaries.NUMBERS_NAMES_IMPLEMENTATION_SOURCES
            )
            semantic.parent.mkdir(parents=True)
            semantic.write_text(
                "\n".join(
                    [
                        "pub fn names(r#source_bytes: &[u8], r#object_id: u64, "
                        "archive_name: &str, component_name: &str, member_name: &str, "
                        "entry_name: &str) -> "
                        "(DocumentArchive, IWorkPackage, SourceCatalog) {}",
                        "pub type Edit = buffa::DocumentArchiveView;",
                        "impl prost::Message for names::Patch {}",
                    ]
                )
                + "\n",
                encoding="utf-8",
            )
            transaction.parent.mkdir(parents=True, exist_ok=True)
            transaction.write_text(
                "\n".join(
                    [
                        "pub type Edit = Archive;",
                        "pub type Patch = ComponentCatalog;",
                        "pub type Commit = EntryEdit;",
                        "pub type Diagnostics = litchi_iwa_core::RawMessage;",
                        "pub type Error = Resolved;",
                        "pub type LimitKind = SnappyStream;",
                        "pub type Path = wire::WireDescent;",
                        "pub type InvalidReason = wire::WireError;",
                        "pub type NamesWireLimits = WireLimits;",
                        "pub type NamesWireResource = WireResourceLimit;",
                        "pub type NamesWireView = WireView;",
                        "pub type NamesNative = NativeObject;",
                    ]
                )
                + "\n",
                encoding="utf-8",
            )
            lib_export, package_export = (
                root / path for path in boundaries.NUMBERS_NAMES_EXPORT_SOURCES
            )
            lib_export.write_text(
                "pub fn edit_names(value: litchi_iwa_protos::GeneratedName) "
                "-> names::Patch {}\n",
                encoding="utf-8",
            )
            package_export.write_text(
                "pub fn apply_names(value: prost_types::MessageInfo) "
                "-> names::Commit {}\n",
                encoding="utf-8",
            )

            violations = boundaries.audit_numbers_names_facade_source_topology(root)

            self.assertEqual(violations, sorted(violations))
            self.assertTrue(
                all(
                    violation.startswith(
                        "focused litchi-numbers names public API exposes "
                    )
                    for violation in violations
                )
            )
            expected_fragments = (
                "raw source bytes source_bytes",
                "raw byte slice &[u8]",
                "raw identifier object_id",
                "physical package name archive_name",
                "physical package name component_name",
                "physical package name member_name",
                "physical package name entry_name",
                "archive/IWA type DocumentArchive",
                "archive/IWA type IWorkPackage",
                "archive/IWA type SourceCatalog",
                "protobuf type buffa",
                "archive/IWA type DocumentArchiveView",
                "protobuf type prost",
                "protobuf type Message",
                "archive/IWA type Archive",
                "archive/IWA type ComponentCatalog",
                "archive/IWA type EntryEdit",
                "archive/IWA type litchi_iwa_core",
                "archive/IWA type RawMessage",
                "archive/IWA type Resolved",
                "archive/IWA type SnappyStream",
                "wire type wire",
                "wire type WireDescent",
                "wire type WireError",
                "wire type WireLimits",
                "wire type WireResourceLimit",
                "wire type WireView",
                "native object NativeObject",
                "archive/IWA type litchi_iwa_protos",
                "generated type GeneratedName",
                "protobuf type prost_types",
                "archive/IWA type MessageInfo",
            )
            self.assertEqual(len(violations), 33)
            for fragment in expected_fragments:
                with self.subTest(fragment=fragment):
                    self.assertTrue(
                        any(fragment in violation for violation in violations),
                        msg=f"missing focused Numbers names leak: {fragment}",
                    )

    def test_focused_numbers_names_api_ignores_nested_private_and_other_owners(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            semantic, transaction = (
                root / path for path in boundaries.NUMBERS_NAMES_IMPLEMENTATION_SOURCES
            )
            semantic.parent.mkdir(parents=True)
            private_aliases = [
                (
                    "struct" if index % 2 == 0 else "pub(crate) struct"
                )
                + f" {name};"
                for index, name in enumerate(sorted(boundaries.NUMBERS_NAMES_FLAT_ALIASES))
            ]
            semantic.write_text(
                "\n".join(
                    [
                        "// pub type NameEdit = DocumentArchive;",
                        'const NOTE: &str = "pub struct NamesPatch;";',
                        "/* pub type SheetNameCommit = buffa::DocumentArchiveView; */",
                        *[f"pub struct {name};" for name in boundaries.NUMBERS_NAMES_CANONICAL_TYPES],
                        *[f"pub struct {name};" for name in boundaries.NUMBERS_NAMES_OPTIONAL_TYPES],
                        "pub struct NamesSnapshot;",
                        "pub struct NameEditor;",
                        "fn source_bytes(source_bytes: &[u8], object_id: u64) "
                        "-> SourceBytes { todo!() }",
                        "pub(crate) fn restricted(archive: DocumentArchive) {}",
                        *private_aliases,
                        "impl NameEdit {}",
                        "impl prost::Message for Unrelated {}",
                    ]
                )
                + "\n",
                encoding="utf-8",
            )
            transaction.parent.mkdir(parents=True, exist_ok=True)
            transaction.write_text(
                "pub fn edit_names(edit: Edit) -> Result<Commit, Error> { todo!() }\n",
                encoding="utf-8",
            )
            lib_export, package_export = (
                root / path for path in boundaries.NUMBERS_NAMES_EXPORT_SOURCES
            )
            restricted_roots = "\n".join(
                (
                    "use" if index % 2 == 0 else "pub(crate) use"
                )
                + f" crate::names::{name};"
                for index, name in enumerate(sorted(boundaries.NUMBERS_NAMES_SHORT_NAMES))
            )
            safe_exports = (
                "pub mod names;\n"
                "// pub use crate::names::{Edit, Patch};\n"
                'const NOTE: &str = "pub use crate::names::*;";\n'
                "pub fn edit_names(edit: names::Edit) "
                "-> Result<names::Commit, names::Error> { todo!() }\n"
                + restricted_roots
                + "\npub(crate) use crate::names::*;\n"
                "pub use crate::render::{Edit, Patch, Commit, Diagnostics, Error, "
                "LimitKind, Path, InvalidReason};\n"
                "pub use crate::render::*;\n"
                "pub fn unrelated(object_id: u64) -> DocumentArchive { todo!() }\n"
            )
            lib_export.write_text(safe_exports, encoding="utf-8")
            package_export.write_text(
                safe_exports.replace("pub mod names;\n", "")
                + "// pub mod names;\n"
                + 'const MODULE_NOTE: &str = "pub mod names;";\n'
                + 'const RAW_MODULE_NOTE: &str = r#"pub mod r#names {}"#;\n'
                + "/* pub mod r#names {} */\n"
                + "mod names;\n"
                + "pub(crate) mod r#names;\n"
                + "pub(super) mod names {}\n"
                + "pub(in crate) mod names;\n",
                encoding="utf-8",
            )
            nonfocused = root / boundaries.NUMBERS_SOURCE_ROOT / "table.rs"
            nonfocused.write_text(
                "\n".join(
                    f"pub struct {name};"
                    for name in sorted(boundaries.NUMBERS_NAMES_FLAT_ALIASES)
                )
                + "\npub fn names(object_id: u64) -> DocumentArchive { todo!() }\n",
                encoding="utf-8",
            )
            other_owner = root / "crates/litchi-pages/src/names.rs"
            other_owner.parent.mkdir(parents=True)
            other_owner.write_text(
                "\n".join(
                    f"pub struct {name};"
                    for name in sorted(boundaries.NUMBERS_NAMES_FLAT_ALIASES)
                )
                + "\npub use crate::names::{Edit, Patch, Commit, Diagnostics, Error, "
                "LimitKind, Path, InvalidReason};\n"
                "pub fn names(object_id: u64) -> DocumentArchive { todo!() }\n",
                encoding="utf-8",
            )

            self.assertEqual(
                boundaries.audit_numbers_names_facade_source_topology(root), []
            )

    def test_focused_numbers_names_public_api_rejects_all_flat_aliases(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            sources = tuple(
                root / path
                for path in (
                    boundaries.NUMBERS_NAMES_IMPLEMENTATION_SOURCES
                    + boundaries.NUMBERS_NAMES_EXPORT_SOURCES
                )
            )
            declarations: dict[Path, list[str]] = {path: [] for path in sources}
            expected: list[str] = []
            relative_sources = (
                boundaries.NUMBERS_NAMES_IMPLEMENTATION_SOURCES
                + boundaries.NUMBERS_NAMES_EXPORT_SOURCES
            )
            for index, name in enumerate(sorted(boundaries.NUMBERS_NAMES_FLAT_ALIASES)):
                source_index = index % len(sources)
                path = sources[source_index]
                if source_index == 0:
                    declaration = f"pub struct {name};"
                elif source_index == 1:
                    declaration = f"pub type {name} = Edit;"
                elif source_index == 2:
                    declaration = f"pub use crate::legacy::Legacy as {name};"
                else:
                    declaration = f"pub use crate::legacy::Legacy as {name};"
                declarations[path].append(declaration)
                expected.append(
                    "focused litchi-numbers names public API retains flat alias "
                    f"{name}: {relative_sources[source_index]}:"
                    f"{len(declarations[path])}"
                )
            for path, lines in declarations.items():
                path.parent.mkdir(parents=True, exist_ok=True)
                path.write_text("\n".join(lines) + "\n", encoding="utf-8")

            self.assertEqual(
                boundaries.audit_numbers_names_facade_source_topology(root),
                sorted(expected),
            )

    def test_focused_numbers_names_exports_reject_root_aliases_and_globs(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            lib_export, package_export = (
                root / path for path in boundaries.NUMBERS_NAMES_EXPORT_SOURCES
            )
            lib_export.parent.mkdir(parents=True)
            lib_export.write_text(
                "pub use crate::names::{Edit, Patch as Patch, Path, InvalidReason};\n"
                "pub type Commit = crate::names::Commit;\n",
                encoding="utf-8",
            )
            package_export.write_text(
                "pub use crate::names::{Diagnostics, Error, LimitKind};\n"
                "pub use crate::names::*;\n",
                encoding="utf-8",
            )

            self.assertEqual(
                boundaries.audit_numbers_names_facade_source_topology(root),
                sorted(
                    [
                        *[
                            "focused litchi-numbers names public API retains root "
                            f"alias {name}: crates/litchi-numbers/src/lib.rs:1"
                            for name in ("Edit", "Patch", "Path", "InvalidReason")
                        ],
                        "focused litchi-numbers names public API retains root alias "
                        "Commit: crates/litchi-numbers/src/lib.rs:2",
                        *[
                            "focused litchi-numbers names public API retains root "
                            f"alias {name}: crates/litchi-numbers/src/package.rs:1"
                            for name in ("Diagnostics", "Error", "LimitKind")
                        ],
                        "focused litchi-numbers names public API retains root aliases "
                        "via names glob: crates/litchi-numbers/src/package.rs:2",
                    ]
                ),
            )

    def test_focused_numbers_names_rejects_duplicate_public_package_module(
        self,
    ) -> None:
        for declaration in (
            "pub mod names;",
            "pub mod r#names;",
            "pub mod names {}",
            "pub\nmod\nr#names\n{}",
        ):
            with self.subTest(declaration=declaration):
                with tempfile.TemporaryDirectory() as directory:
                    root = Path(directory)
                    package_export = (
                        root / boundaries.NUMBERS_NAMES_EXPORT_SOURCES[1]
                    )
                    package_export.parent.mkdir(parents=True)
                    package_export.write_text(declaration + "\n", encoding="utf-8")

                    self.assertEqual(
                        boundaries.audit_numbers_names_facade_source_topology(root),
                        [
                            "focused litchi-numbers names public API exposes "
                            "duplicate package::names module: "
                            "crates/litchi-numbers/src/package.rs:1"
                        ],
                    )

    def test_numbers_table_header_settings_boundary_inventories_are_exact(
        self,
    ) -> None:
        self.assertEqual(
            boundaries.RETIRED_IWA_NUMBERS_TABLE_HEADER_SETTINGS_METHODS,
            ("table_header_settings", "set_table_header_settings"),
        )
        self.assertEqual(
            boundaries.RETIRED_IWA_NUMBERS_TABLE_HEADER_SETTINGS_EXAMPLE,
            Path("crates/litchi-iwa/examples/edit_numbers_table_headers.rs"),
        )
        self.assertEqual(
            boundaries.NUMBERS_TABLE_HEADER_SETTINGS_IMPLEMENTATION_SOURCES,
            (
                Path("crates/litchi-numbers/src/table/headers.rs"),
                Path("crates/litchi-numbers/src/table/headers/transaction.rs"),
                Path("crates/litchi-numbers/src/package/table_headers.rs"),
                Path("crates/litchi-numbers/src/package/table_headers/api.rs"),
                Path("crates/litchi-numbers/src/package/table_headers/dependencies.rs"),
                Path("crates/litchi-numbers/src/package/table_headers/error.rs"),
                Path("crates/litchi-numbers/src/package/table_headers/ownership.rs"),
                Path("crates/litchi-numbers/src/package/table_headers/resolve.rs"),
                Path("crates/litchi-numbers/src/package/table_headers/rewrite.rs"),
            ),
        )
        self.assertEqual(
            boundaries.NUMBERS_TABLE_HEADER_SETTINGS_EXPORT_SOURCES,
            (
                Path("crates/litchi-numbers/src/lib.rs"),
                Path("crates/litchi-numbers/src/package.rs"),
                Path("crates/litchi-numbers/src/table.rs"),
            ),
        )
        self.assertEqual(
            boundaries.NUMBERS_TABLE_HEADER_SETTINGS_CANONICAL_TYPES,
            (
                "Edit",
                "Patch",
                "Commit",
                "Diagnostics",
                "Error",
                "LimitKind",
                "Path",
                "InvalidReason",
            ),
        )
        self.assertEqual(
            boundaries.NUMBERS_TABLE_HEADER_SETTINGS_PACKAGE_METHODS,
            ("table_header_settings", "edit_table_headers", "apply_table_headers"),
        )
        self.assertEqual(
            boundaries.NUMBERS_TABLE_HEADER_SETTINGS_SEMANTIC_TYPES,
            ("Count", "Error", "Settings"),
        )
        self.assertEqual(
            boundaries.NUMBERS_TABLE_HEADER_SETTINGS_FLAT_ALIAS_PREFIXES,
            ("HeaderSettings", "TableHeader", "TableHeaders", "TableHeaderSettings"),
        )
        self.assertEqual(
            boundaries.NUMBERS_TABLE_HEADER_SETTINGS_FLAT_ALIASES,
            frozenset(
                prefix + suffix
                for prefix in (
                    "HeaderSettings",
                    "TableHeader",
                    "TableHeaders",
                    "TableHeaderSettings",
                )
                for suffix in (
                    "Edit",
                    "Patch",
                    "Commit",
                    "Diagnostics",
                    "Error",
                    "LimitKind",
                    "Path",
                    "InvalidReason",
                )
            ),
        )
        self.assertEqual(
            boundaries.NUMBERS_TABLE_HEADER_SETTINGS_FLAT_SEMANTIC_ALIASES,
            frozenset(
                {
                    "HeaderCount",
                    "HeaderSettings",
                    "TableHeaderCount",
                    "TableHeaderSettings",
                }
            ),
        )
        self.assertEqual(
            boundaries.NUMBERS_TABLE_HEADER_SETTINGS_ROOT_ALIASES,
            frozenset(
                boundaries.NUMBERS_TABLE_HEADER_SETTINGS_CANONICAL_TYPES
                + boundaries.NUMBERS_TABLE_HEADER_SETTINGS_SEMANTIC_TYPES
            ),
        )
        self.assertEqual(
            boundaries.NUMBERS_TABLE_HEADER_SETTINGS_PHYSICAL_TYPES,
            frozenset(
                {
                    "Archive",
                    "ComponentCatalog",
                    "EntryEdit",
                    "ExactArtifacts",
                    "IWorkPackage",
                    "PhysicalSource",
                    "RawMessage",
                    "Resolved",
                    "SnappyStream",
                    "SourceCatalog",
                    "TableHeaderSettingsSnapshot",
                    "TableInfoSnapshot",
                }
            ),
        )
        self.assertEqual(
            boundaries.NUMBERS_TABLE_HEADER_SETTINGS_WIRE_TYPES,
            frozenset(
                {
                    "DecodeOptions",
                    "NestedFieldEdit",
                    "NestedFieldReplacement",
                    "WireDescent",
                    "WireError",
                    "WireLimits",
                    "WireResourceLimit",
                    "WireView",
                }
            ),
        )

    def test_retired_iwa_numbers_table_header_settings_surface_cannot_return(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            host = root / boundaries.IWA_NUMBERS_SOURCE_ROOT / "editor/table_headers.rs"
            host.parent.mkdir(parents=True)
            host.write_text(
                "fn r#table_header_settings() {}\n"
                "pub(in crate::numbers) async unsafe fn set_table_header_settings() {}\n",
                encoding="utf-8",
            )
            example = root / boundaries.RETIRED_IWA_NUMBERS_TABLE_HEADER_SETTINGS_EXAMPLE
            example.parent.mkdir(parents=True, exist_ok=True)
            example.write_text("// retired example returned\n", encoding="utf-8")

            self.assertEqual(
                boundaries.audit_iwa_numbers_table_header_settings_source_topology(
                    root
                ),
                sorted(
                    [
                        "retired litchi-iwa Numbers table-header settings method "
                        "table_header_settings: "
                        "crates/litchi-iwa/src/numbers/editor/table_headers.rs:1",
                        "retired litchi-iwa Numbers table-header settings method "
                        "set_table_header_settings: "
                        "crates/litchi-iwa/src/numbers/editor/table_headers.rs:2",
                        "retired litchi-iwa Numbers table-header settings example "
                        "returned: "
                        "crates/litchi-iwa/examples/edit_numbers_table_headers.rs",
                    ]
                ),
            )

    def test_retired_iwa_numbers_table_header_settings_readme_calls_and_example(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            readme = root / boundaries.IWA_NUMBERS_README
            readme.parent.mkdir(parents=True)
            readme.write_text(
                "\n".join(
                    [
                        "let settings = numbers",
                        "    .",
                        "    r#table_header_settings",
                        "    (",
                        "numbers_editor.r#set_table_header_settings(",
                        "crate::numbers::NumbersEditor",
                        "    ::",
                        "    set_table_header_settings(",
                        "r#NumbersEditor::r#table_header_settings(",
                        "Run `edit_numbers_table_headers`.",
                        "cargo run --example edit_numbers_table_headers.rs",
                    ]
                )
                + "\n",
                encoding="utf-8",
            )

            self.assertEqual(
                boundaries.audit_iwa_numbers_table_header_settings_source_topology(
                    root
                ),
                sorted(
                    [
                        "retired litchi-iwa Numbers table-header settings README call "
                        "table_header_settings: crates/litchi-iwa/README.md:3",
                        "retired litchi-iwa Numbers table-header settings README call "
                        "set_table_header_settings: crates/litchi-iwa/README.md:5",
                        "retired litchi-iwa Numbers table-header settings README call "
                        "set_table_header_settings: crates/litchi-iwa/README.md:8",
                        "retired litchi-iwa Numbers table-header settings README call "
                        "table_header_settings: crates/litchi-iwa/README.md:9",
                        "retired litchi-iwa Numbers table-header settings README "
                        "example reference edit_numbers_table_headers: "
                        "crates/litchi-iwa/README.md:10",
                        "retired litchi-iwa Numbers table-header settings README "
                        "example reference edit_numbers_table_headers: "
                        "crates/litchi-iwa/README.md:11",
                    ]
                ),
            )

    def test_iwa_numbers_table_header_settings_policy_ignores_safe_helpers_and_owners(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            host = root / boundaries.IWA_NUMBERS_SOURCE_ROOT / "editor/table_headers.rs"
            host.parent.mkdir(parents=True)
            host.write_text(
                "\n".join(
                    [
                        "// pub fn table_header_settings() {}",
                        'const NOTE: &str = "fn set_table_header_settings() {}";',
                        "/* fn table_header_settings() {}",
                        "   /* fn set_table_header_settings() {} */",
                        "   fn table_header_settings() {} */",
                        'const RAW_NOTE: &str = r###"fn set_table_header_settings() {}"###;',
                        "pub(super) fn read_table_header_settings() {}",
                        "pub(super) fn read_attached_table_header_settings() {}",
                        "pub(super) fn set_attached_table_header_settings() {}",
                        "pub(crate) fn table_header_settings_in_package() {}",
                        "pub(crate) fn set_table_header_settings_in_package() {}",
                        "pub(super) fn read_table_header_settings_wire() {}",
                        "pub(super) fn write_table_header_settings_wire() {}",
                        "pub fn table_header_settings_snapshot() {}",
                        "pub fn set_table_header_settings_for_model() {}",
                    ]
                )
                + "\n",
                encoding="utf-8",
            )
            readme = root / boundaries.IWA_NUMBERS_README
            readme.write_text(
                "\n".join(
                    [
                        "The table_header_settings names are retired.",
                        "Use `table_header_settings` rather than `.table_header_settings`.",
                        "table_header_settings(",
                        "set_table_header_settings(",
                        "numbers.table_header_settings_snapshot(",
                        "editor.set_table_header_settings_for_model(",
                        "editor.table_header_settings(",
                        "editor.set_table_header_settings(",
                        "pages.table_header_settings(",
                        "keynote.set_table_header_settings(",
                        "package.table_header_settings(",
                        "package.edit_table_headers(",
                        "edit_numbers_table_header.rs",
                        "edit_numbers_table_headers_old.rs",
                    ]
                )
                + "\n",
                encoding="utf-8",
            )
            declarations = (
                "pub fn table_header_settings() {}\n"
                "pub fn set_table_header_settings() {}\n"
            )
            for relative in (
                Path("crates/litchi-numbers/src/package/table_headers.rs"),
                Path("crates/litchi-iwa/src/pages/editor/table_headers.rs"),
                Path("crates/litchi-iwa/src/keynote/editor/table_headers.rs"),
            ):
                path = root / relative
                path.parent.mkdir(parents=True, exist_ok=True)
                path.write_text(declarations, encoding="utf-8")
            non_rust = root / boundaries.IWA_NUMBERS_SOURCE_ROOT / "table_headers.txt"
            non_rust.write_text(declarations, encoding="utf-8")
            near_example = (
                root
                / boundaries.RETIRED_IWA_NUMBERS_TABLE_HEADER_SETTINGS_EXAMPLE.with_name(
                    "edit_numbers_table_headers_old.rs"
                )
            )
            near_example.parent.mkdir(parents=True, exist_ok=True)
            near_example.write_text("// near example\n", encoding="utf-8")
            for relative in (Path("README.md"), Path("crates/litchi-numbers/README.md")):
                other = root / relative
                other.parent.mkdir(parents=True, exist_ok=True)
                other.write_text(
                    "numbers.table_header_settings(\nedit_numbers_table_headers\n",
                    encoding="utf-8",
                )

            self.assertEqual(
                boundaries.audit_iwa_numbers_table_header_settings_source_topology(
                    root
                ),
                [],
            )
    def test_focused_numbers_table_header_settings_requires_canonical_types(
        self,
    ) -> None:
        for missing in boundaries.NUMBERS_TABLE_HEADER_SETTINGS_CANONICAL_TYPES:
            with self.subTest(missing=missing):
                with tempfile.TemporaryDirectory() as directory:
                    root = Path(directory)
                    semantic = (
                        root / boundaries.NUMBERS_TABLE_HEADER_SETTINGS_SEMANTIC_SOURCE
                    )
                    semantic.parent.mkdir(parents=True)
                    semantic.write_text("pub mod transaction;\n", encoding="utf-8")
                    transaction = (
                        root
                        / boundaries.NUMBERS_TABLE_HEADER_SETTINGS_TRANSACTION_SOURCE
                    )
                    transaction.parent.mkdir(parents=True, exist_ok=True)
                    transaction.write_text(
                        "".join(
                            f"pub struct {name};\n"
                            for name in boundaries.NUMBERS_TABLE_HEADER_SETTINGS_CANONICAL_TYPES
                            if name != missing
                        ),
                        encoding="utf-8",
                    )
                    table_export = (
                        root / boundaries.NUMBERS_TABLE_HEADER_SETTINGS_EXPORT_SOURCES[2]
                    )
                    table_export.write_text("pub mod headers;\n", encoding="utf-8")
                    lib_export = (
                        root / boundaries.NUMBERS_TABLE_HEADER_SETTINGS_EXPORT_SOURCES[0]
                    )
                    lib_export.write_text("pub mod table;\n", encoding="utf-8")

                    self.assertEqual(
                        boundaries.audit_numbers_table_header_settings_facade_source_topology(
                            root
                        ),
                        [
                            "focused litchi-numbers table-header settings public API "
                            f"is missing canonical transaction type {missing}: "
                            "crates/litchi-numbers/src/table/headers.rs"
                        ],
                    )

    def test_focused_numbers_table_header_settings_module_visibility(self) -> None:
        missing_cases = (
            (
                boundaries.NUMBERS_TABLE_HEADER_SETTINGS_SEMANTIC_SOURCE,
                (
                    "",
                    "mod transaction;\n",
                    "pub(crate) mod transaction;\n",
                    "pub(super) mod r#transaction {}\n",
                    "// pub mod transaction;\n",
                    'const NOTE: &str = "pub mod transaction;";\n',
                ),
                "focused litchi-numbers table-header settings public API is missing "
                "canonical headers::transaction module: "
                "crates/litchi-numbers/src/table/headers.rs",
            ),
            (
                boundaries.NUMBERS_TABLE_HEADER_SETTINGS_EXPORT_SOURCES[2],
                (
                    "",
                    "mod headers;\n",
                    "pub(crate) mod headers;\n",
                    "pub(in crate) mod r#headers {}\n",
                    "/* pub mod headers; */\n",
                ),
                "focused litchi-numbers table-header settings public API is missing "
                "canonical table::headers module: crates/litchi-numbers/src/table.rs",
            ),
            (
                boundaries.NUMBERS_TABLE_HEADER_SETTINGS_EXPORT_SOURCES[0],
                (
                    "",
                    "mod table;\n",
                    "pub(crate) mod table;\n",
                    "pub(super) mod r#table {}\n",
                    'const RAW_NOTE: &str = r#"pub mod table;"#;\n',
                ),
                "focused litchi-numbers table-header settings public API is missing "
                "canonical root table module: crates/litchi-numbers/src/lib.rs",
            ),
        )
        for relative, declarations, diagnostic in missing_cases:
            for declaration in declarations:
                with self.subTest(relative=relative, declaration=declaration):
                    with tempfile.TemporaryDirectory() as directory:
                        root = Path(directory)
                        add_numbers_table_header_settings_canonical_scaffold(root)
                        path = root / relative
                        path.write_text(declaration, encoding="utf-8")

                        self.assertEqual(
                            boundaries.audit_numbers_table_header_settings_facade_source_topology(
                                root
                            ),
                            [diagnostic],
                        )

        accepted_cases = (
            (
                boundaries.NUMBERS_TABLE_HEADER_SETTINGS_SEMANTIC_SOURCE,
                (
                    "pub mod transaction;\n",
                    "pub mod r#transaction;\n",
                    "pub\nmod\nr#transaction\n;\n",
                ),
            ),
            (
                boundaries.NUMBERS_TABLE_HEADER_SETTINGS_EXPORT_SOURCES[2],
                (
                    "pub mod headers;\n",
                    "pub mod r#headers;\n",
                    "pub mod headers {}\n",
                    "pub\nmod\nr#headers\n{}\n",
                ),
            ),
            (
                boundaries.NUMBERS_TABLE_HEADER_SETTINGS_EXPORT_SOURCES[0],
                (
                    "pub mod table;\n",
                    "pub mod r#table;\n",
                    "pub mod table {}\n",
                    "pub\nmod\nr#table\n{}\n",
                ),
            ),
        )
        for relative, declarations in accepted_cases:
            for declaration in declarations:
                with self.subTest(relative=relative, declaration=declaration):
                    with tempfile.TemporaryDirectory() as directory:
                        root = Path(directory)
                        add_numbers_table_header_settings_canonical_scaffold(root)
                        path = root / relative
                        path.write_text(declaration, encoding="utf-8")

                        self.assertEqual(
                            boundaries.audit_numbers_table_header_settings_facade_source_topology(
                                root
                            ),
                            [],
                        )

    def test_focused_numbers_table_header_settings_accepts_inline_transaction(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            semantic = root / boundaries.NUMBERS_TABLE_HEADER_SETTINGS_SEMANTIC_SOURCE
            semantic.parent.mkdir(parents=True)
            semantic.write_text(
                "pub mod transaction {\n"
                + "".join(
                    f"    pub struct {name};\n"
                    for name in boundaries.NUMBERS_TABLE_HEADER_SETTINGS_CANONICAL_TYPES
                )
                + "}\n",
                encoding="utf-8",
            )
            table_export = (
                root / boundaries.NUMBERS_TABLE_HEADER_SETTINGS_EXPORT_SOURCES[2]
            )
            table_export.write_text("pub mod headers;\n", encoding="utf-8")
            lib_export = (
                root / boundaries.NUMBERS_TABLE_HEADER_SETTINGS_EXPORT_SOURCES[0]
            )
            lib_export.write_text("pub mod table;\n", encoding="utf-8")

            self.assertEqual(
                boundaries.audit_numbers_table_header_settings_facade_source_topology(
                    root
                ),
                [],
            )

    def test_focused_numbers_table_header_settings_rejects_physical_leaks(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            semantic = root / boundaries.NUMBERS_TABLE_HEADER_SETTINGS_SEMANTIC_SOURCE
            transaction = (
                root / boundaries.NUMBERS_TABLE_HEADER_SETTINGS_TRANSACTION_SOURCE
            )
            owner = root / boundaries.NUMBERS_TABLE_HEADER_SETTINGS_OWNER_SOURCE
            semantic.parent.mkdir(parents=True)
            semantic.write_text(
                "pub fn headers(r#source_bytes: &[u8], r#object_id: u64, "
                "archive_name: &str, component_name: &str, member_name: &str, "
                "entry_name: &str) -> "
                "(DocumentArchive, IWorkPackage, SourceCatalog) {}\n"
                "pub type Count = buffa::DocumentArchiveView;\n"
                "impl prost::Message for headers::Settings {}\n",
                encoding="utf-8",
            )
            transaction.parent.mkdir(parents=True, exist_ok=True)
            transaction.write_text(
                "\n".join(
                    [
                        "pub type Edit = Archive;",
                        "pub type Patch = ComponentCatalog;",
                        "pub type Commit = EntryEdit;",
                        "pub type Diagnostics = ExactArtifacts;",
                        "pub type Error = IWorkPackage;",
                        "pub type LimitKind = PhysicalSource;",
                        "pub type Path = RawMessage;",
                        "pub type InvalidReason = Resolved;",
                        "pub type HeaderSnappy = SnappyStream;",
                        "pub type HeaderCatalog = SourceCatalog;",
                        "pub type HeaderSnapshot = TableHeaderSettingsSnapshot;",
                        "pub type HeaderInfo = TableInfoSnapshot;",
                        "pub type HeaderDecode = DecodeOptions;",
                        "pub type HeaderFieldEdit = NestedFieldEdit;",
                        "pub type HeaderReplacement = NestedFieldReplacement;",
                        "pub type HeaderDescent = wire::WireDescent;",
                        "pub type HeaderWireError = wire::WireError;",
                        "pub type HeaderWireLimits = WireLimits;",
                        "pub type HeaderWireResource = WireResourceLimit;",
                        "pub type HeaderWireView = WireView;",
                        "pub type HeaderNative = NativeObject;",
                    ]
                )
                + "\n",
                encoding="utf-8",
            )
            owner.parent.mkdir(parents=True, exist_ok=True)
            owner.write_text(
                "pub type HeaderOwner = litchi_iwa_core::RawObject;\n",
                encoding="utf-8",
            )
            lib_export, package_export, table_export = (
                root / path
                for path in boundaries.NUMBERS_TABLE_HEADER_SETTINGS_EXPORT_SOURCES
            )
            lib_export.write_text(
                "pub fn edit_table_headers("
                "value: litchi_iwa_protos::GeneratedHeaderSettings) "
                "-> table::headers::transaction::Edit {}\n",
                encoding="utf-8",
            )
            package_export.write_text(
                "pub fn apply_table_headers(value: prost_types::MessageInfo) "
                "-> table::headers::transaction::Commit {}\n",
                encoding="utf-8",
            )
            table_export.write_text(
                "pub fn table_header_settings(value: SourceBytes) "
                "-> headers::Settings {}\n",
                encoding="utf-8",
            )
            add_numbers_table_header_settings_canonical_scaffold(root)

            violations = (
                boundaries.audit_numbers_table_header_settings_facade_source_topology(
                    root
                )
            )

            self.assertEqual(violations, sorted(violations))
            self.assertTrue(
                all(
                    violation.startswith(
                        "focused litchi-numbers table-header settings public API "
                        "exposes "
                    )
                    for violation in violations
                )
            )
            expected_fragments = (
                "raw source bytes source_bytes",
                "raw byte slice &[u8]",
                "raw identifier object_id",
                "physical package name archive_name",
                "physical package name component_name",
                "physical package name member_name",
                "physical package name entry_name",
                "archive/IWA type DocumentArchive",
                "archive/IWA type IWorkPackage",
                "archive/IWA type SourceCatalog",
                "protobuf type buffa",
                "archive/IWA type DocumentArchiveView",
                "protobuf type prost",
                "protobuf type Message",
                "archive/IWA type Archive",
                "archive/IWA type ComponentCatalog",
                "archive/IWA type EntryEdit",
                "archive/IWA type ExactArtifacts",
                "archive/IWA type PhysicalSource",
                "archive/IWA type RawMessage",
                "archive/IWA type Resolved",
                "archive/IWA type SnappyStream",
                "archive/IWA type TableHeaderSettingsSnapshot",
                "archive/IWA type TableInfoSnapshot",
                "wire type DecodeOptions",
                "wire type NestedFieldEdit",
                "wire type NestedFieldReplacement",
                "wire type wire",
                "wire type WireDescent",
                "wire type WireError",
                "wire type WireLimits",
                "wire type WireResourceLimit",
                "wire type WireView",
                "native object NativeObject",
                "archive/IWA type litchi_iwa_core",
                "native object RawObject",
                "archive/IWA type litchi_iwa_protos",
                "generated type GeneratedHeaderSettings",
                "protobuf type prost_types",
                "archive/IWA type MessageInfo",
                "raw source bytes SourceBytes",
            )
            self.assertEqual(len(violations), 44)
            for fragment in expected_fragments:
                with self.subTest(fragment=fragment):
                    self.assertTrue(
                        any(fragment in violation for violation in violations),
                        msg=f"missing focused table-header leak: {fragment}",
                    )

    def test_focused_numbers_table_header_settings_scans_split_owner_modules(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            add_numbers_table_header_settings_canonical_scaffold(root)
            split_owner = root / Path(
                "crates/litchi-numbers/src/package/table_headers/api.rs"
            )
            split_owner.parent.mkdir(parents=True, exist_ok=True)
            split_owner.write_text(
                "impl Package {\n"
                "    pub fn table_header_settings(&self, object_id: u64) "
                "-> Archive { todo!() }\n"
                "}\n",
                encoding="utf-8",
            )

            violations = (
                boundaries.audit_numbers_table_header_settings_facade_source_topology(
                    root
                )
            )

            self.assertTrue(
                any(
                    "raw identifier object_id" in violation
                    and "package/table_headers/api.rs" in violation
                    for violation in violations
                )
            )
            self.assertTrue(
                any(
                    "archive/IWA type Archive" in violation
                    and "package/table_headers/api.rs" in violation
                    for violation in violations
                )
            )

    def test_focused_numbers_table_header_settings_allows_nested_private_surfaces(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            semantic = root / boundaries.NUMBERS_TABLE_HEADER_SETTINGS_SEMANTIC_SOURCE
            transaction = (
                root / boundaries.NUMBERS_TABLE_HEADER_SETTINGS_TRANSACTION_SOURCE
            )
            owner = root / boundaries.NUMBERS_TABLE_HEADER_SETTINGS_OWNER_SOURCE
            semantic.parent.mkdir(parents=True)
            semantic.write_text(
                "pub struct Count;\n"
                "pub struct Settings;\n"
                "pub enum Error {}\n"
                "// pub struct HeaderSettings;\n"
                'const NOTE: &str = "pub type TableHeaderSettingsEdit = Archive";\n'
                "/* pub struct TableHeaderCount; */\n"
                "fn source_bytes(source_bytes: &[u8], object_id: u64) "
                "-> SourceBytes { todo!() }\n"
                "pub(crate) fn restricted(archive: DocumentArchive) {}\n",
                encoding="utf-8",
            )
            all_flat_aliases = sorted(
                boundaries.NUMBERS_TABLE_HEADER_SETTINGS_FLAT_ALIASES
                | boundaries.NUMBERS_TABLE_HEADER_SETTINGS_FLAT_SEMANTIC_ALIASES
            )
            transaction.parent.mkdir(parents=True, exist_ok=True)
            transaction.write_text(
                "\n".join(
                    (
                        "struct" if index % 2 == 0 else "pub(crate) struct"
                    )
                    + f" {name};"
                    for index, name in enumerate(all_flat_aliases)
                )
                + "\nimpl TableHeaderSettingsEdit {}\n",
                encoding="utf-8",
            )
            owner.parent.mkdir(parents=True, exist_ok=True)
            owner.write_text(
                "pub fn table_header_settings() -> table::headers::Settings { todo!() }\n"
                "pub fn edit_table_headers() -> table::headers::transaction::Edit "
                "{ todo!() }\n"
                "pub fn apply_table_headers() -> table::headers::transaction::Commit "
                "{ todo!() }\n",
                encoding="utf-8",
            )
            lib_export, package_export, table_export = (
                root / path
                for path in boundaries.NUMBERS_TABLE_HEADER_SETTINGS_EXPORT_SOURCES
            )
            private_roots = "\n".join(
                ("use" if index % 2 == 0 else "pub(crate) use")
                + f" crate::table::headers::{name};"
                for index, name in enumerate(
                    sorted(boundaries.NUMBERS_TABLE_HEADER_SETTINGS_ROOT_ALIASES)
                )
            )
            safe_exports = (
                "// pub use crate::table::headers::{Settings, Edit};\n"
                'const NOTE: &str = "pub use crate::table::headers::*;";\n'
                "pub fn table_header_settings() -> table::headers::Settings { todo!() }\n"
                + private_roots
                + "\npub(crate) use crate::table::headers::*;\n"
                "pub use crate::layout::{Count, Settings, Edit, Patch, Commit, "
                "Diagnostics, Error, LimitKind, Path, InvalidReason};\n"
                "pub use crate::layout::*;\n"
                "pub fn unrelated(object_id: u64) -> DocumentArchive { todo!() }\n"
            )
            lib_export.write_text(safe_exports, encoding="utf-8")
            table_export.write_text(safe_exports, encoding="utf-8")
            package_export.write_text(
                safe_exports
                + "// pub mod table_headers;\n"
                + 'const MODULE_NOTE: &str = "pub mod table_headers;";\n'
                + 'const RAW_NOTE: &str = r#"pub mod r#table_headers {}"#;\n'
                + "/* pub mod r#table_headers {} */\n"
                + "mod table_headers;\n"
                + "pub(crate) mod r#table_headers;\n"
                + "pub(super) mod table_headers {}\n"
                + "pub(in crate) mod table_headers;\n",
                encoding="utf-8",
            )
            nonfocused = root / boundaries.NUMBERS_SOURCE_ROOT / "cell.rs"
            nonfocused.write_text(
                "\n".join(f"pub struct {name};" for name in all_flat_aliases)
                + "\npub fn table_headers(object_id: u64) -> DocumentArchive { todo!() }\n",
                encoding="utf-8",
            )
            other_owner = root / "crates/litchi-pages/src/table_headers.rs"
            other_owner.parent.mkdir(parents=True)
            other_owner.write_text(
                "\n".join(f"pub struct {name};" for name in all_flat_aliases)
                + "\npub use crate::table::headers::{Count, Settings, Edit, Patch};\n"
                "pub fn table_headers(object_id: u64) -> DocumentArchive { todo!() }\n",
                encoding="utf-8",
            )
            add_numbers_table_header_settings_canonical_scaffold(root)

            self.assertEqual(
                boundaries.audit_numbers_table_header_settings_facade_source_topology(
                    root
                ),
                [],
            )

    def test_focused_numbers_table_header_settings_rejects_all_flat_aliases(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            relative_sources = (
                boundaries.NUMBERS_TABLE_HEADER_SETTINGS_IMPLEMENTATION_SOURCES
                + boundaries.NUMBERS_TABLE_HEADER_SETTINGS_EXPORT_SOURCES
            )
            sources = tuple(root / path for path in relative_sources)
            declarations: dict[Path, list[str]] = {path: [] for path in sources}
            expected: list[str] = []
            aliases = sorted(
                boundaries.NUMBERS_TABLE_HEADER_SETTINGS_FLAT_ALIASES
                | boundaries.NUMBERS_TABLE_HEADER_SETTINGS_FLAT_SEMANTIC_ALIASES
            )
            for index, name in enumerate(aliases):
                source_index = index % len(sources)
                path = sources[source_index]
                if source_index == 0:
                    declaration = f"pub struct {name};"
                elif source_index == 1:
                    declaration = f"pub type {name} = Edit;"
                elif source_index == 2:
                    declaration = f"pub enum {name} {{ Legacy }}"
                else:
                    declaration = f"pub use crate::legacy::Legacy as {name};"
                declarations[path].append(declaration)
                expected.append(
                    "focused litchi-numbers table-header settings public API retains "
                    f"flat alias {name}: {relative_sources[source_index]}:"
                    f"{len(declarations[path])}"
                )
            for path, lines in declarations.items():
                path.parent.mkdir(parents=True, exist_ok=True)
                path.write_text("\n".join(lines) + "\n", encoding="utf-8")
            add_numbers_table_header_settings_canonical_scaffold(root)

            self.assertEqual(
                boundaries.audit_numbers_table_header_settings_facade_source_topology(
                    root
                ),
                sorted(expected),
            )

    def test_focused_numbers_table_header_settings_exports_reject_root_aliases(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            lib_export, package_export, table_export = (
                root / path
                for path in boundaries.NUMBERS_TABLE_HEADER_SETTINGS_EXPORT_SOURCES
            )
            lib_export.parent.mkdir(parents=True)
            aliases = sorted(boundaries.NUMBERS_TABLE_HEADER_SETTINGS_ROOT_ALIASES)
            first = aliases[:3]
            second = aliases[3:6]
            third = aliases[6:]
            lib_export.write_text(
                "pub use crate::table::headers::{" + ", ".join(first) + "};\n",
                encoding="utf-8",
            )
            package_export.write_text(
                "pub use crate::table_headers::{" + ", ".join(second) + "};\n"
                "pub use crate::table_headers::*;\n",
                encoding="utf-8",
            )
            table_export.write_text(
                "pub use self::headers::{" + ", ".join(third) + "};\n",
                encoding="utf-8",
            )
            add_numbers_table_header_settings_canonical_scaffold(root)

            self.assertEqual(
                boundaries.audit_numbers_table_header_settings_facade_source_topology(
                    root
                ),
                sorted(
                    [
                        *[
                            "focused litchi-numbers table-header settings public API "
                            f"retains root alias {name}: "
                            "crates/litchi-numbers/src/lib.rs:1"
                            for name in first
                        ],
                        *[
                            "focused litchi-numbers table-header settings public API "
                            f"retains root alias {name}: "
                            "crates/litchi-numbers/src/package.rs:1"
                            for name in second
                        ],
                        *[
                            "focused litchi-numbers table-header settings public API "
                            f"retains root alias {name}: "
                            "crates/litchi-numbers/src/table.rs:1"
                            for name in third
                        ],
                        "focused litchi-numbers table-header settings public API "
                        "retains root aliases via table-header glob: "
                        "crates/litchi-numbers/src/package.rs:2",
                    ]
                ),
            )

    def test_focused_numbers_table_header_settings_rejects_public_package_module(
        self,
    ) -> None:
        for declaration in (
            "pub mod table_headers;",
            "pub mod r#table_headers;",
            "pub mod table_headers {}",
            "pub\nmod\nr#table_headers\n{}",
        ):
            with self.subTest(declaration=declaration):
                with tempfile.TemporaryDirectory() as directory:
                    root = Path(directory)
                    package_export = (
                        root / boundaries.NUMBERS_TABLE_HEADER_SETTINGS_EXPORT_SOURCES[1]
                    )
                    package_export.parent.mkdir(parents=True)
                    package_export.write_text(declaration + "\n", encoding="utf-8")
                    add_numbers_table_header_settings_canonical_scaffold(root)

                    self.assertEqual(
                        boundaries.audit_numbers_table_header_settings_facade_source_topology(
                            root
                        ),
                        [
                            "focused litchi-numbers table-header settings public API "
                            "exposes duplicate package::table_headers module: "
                            "crates/litchi-numbers/src/package.rs:1"
                        ],
                    )

    def test_retired_iwa_numbers_document_reader_inventory_is_exact(self) -> None:
        self.assertEqual(
            boundaries.RETIRED_IWA_NUMBERS_DOCUMENT_SOURCE,
            Path("crates/litchi-iwa/src/numbers/document.rs"),
        )
        self.assertEqual(
            boundaries.RETIRED_IWA_NUMBERS_DOCUMENT_TYPES,
            ("NumbersDocument", "NumbersDocumentState", "NumbersDocumentStats"),
        )
        self.assertEqual(
            boundaries.RETIRED_IWA_NUMBERS_SHEET_SOURCE,
            Path("crates/litchi-iwa/src/numbers/sheet.rs"),
        )
        self.assertEqual(
            boundaries.RETIRED_IWA_NUMBERS_SHEET_TYPES, ("NumbersSheet",)
        )
        self.assertEqual(
            boundaries.IWA_NUMBERS_MODULE_SOURCE,
            Path("crates/litchi-iwa/src/numbers/mod.rs"),
        )
        self.assertEqual(
            boundaries.IWA_NUMBERS_FOCUSED_READER_TYPES,
            frozenset({"Document", "Package"}),
        )

    def test_retired_iwa_numbers_document_reader_surface_cannot_return(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            retired = root / boundaries.RETIRED_IWA_NUMBERS_DOCUMENT_SOURCE
            retired.parent.mkdir(parents=True)
            retired.write_text("// retired reader returned\n", encoding="utf-8")
            retired_sheet = root / boundaries.RETIRED_IWA_NUMBERS_SHEET_SOURCE
            retired_sheet.write_text("// retired sheet returned\n", encoding="utf-8")
            module = root / boundaries.IWA_NUMBERS_MODULE_SOURCE
            module.write_text(
                "pub mod r#document;\n"
                "pub(crate) mod sheet;\n"
                "pub use self::r#document::*;\n"
                "pub use sheet::NumbersSheet;\n",
                encoding="utf-8",
            )
            caller = root / "crates/litchi-iwa/src/legacy_numbers.rs"
            caller.parent.mkdir(parents=True, exist_ok=True)
            caller.write_text(
                "pub fn open() -> NumbersDocument { todo!() }\n"
                "pub type State = NumbersDocumentState;\n"
                "pub fn sheets() -> Vec<NumbersSheet> { todo!() }\n"
                "/// Do not restore `NumbersDocumentStats`.\n",
                encoding="utf-8",
            )
            example = root / "crates/litchi-iwa/examples/read_numbers.rs"
            example.parent.mkdir(parents=True)
            example.write_text(
                "use litchi_iwa::numbers::NumbersDocument;\n",
                encoding="utf-8",
            )
            readme = root / boundaries.IWA_NUMBERS_README
            readme.write_text(
                "Never restore `NumbersDocumentStats`.\n", encoding="utf-8"
            )

            violations = boundaries.audit_iwa_numbers_document_source_topology(root)
            joined = "\n".join(violations)
            self.assertIn("document reader source returned", joined)
            self.assertIn("sheet reader source returned", joined)
            self.assertIn("reader module document", joined)
            self.assertIn("reader module sheet", joined)
            self.assertIn("reader local re-export document", joined)
            self.assertIn("reader local re-export sheet", joined)
            self.assertIn("workspace public name NumbersDocument", joined)
            self.assertIn("workspace public name NumbersDocumentState", joined)
            self.assertIn("workspace public name NumbersSheet", joined)
            self.assertIn("workspace public rustdoc NumbersDocumentStats", joined)
            self.assertIn("workspace type usage NumbersDocument", joined)
            self.assertIn("document reader README reference NumbersDocumentStats", joined)

    def test_retired_iwa_numbers_document_module_and_reexport_variants(self) -> None:
        declarations = (
            "mod document;\n",
            "pub(crate) mod r#document {}\n",
            "pub\nmod\ndocument\n{}\n",
            "pub use document::NumbersDocument;\n",
            "pub(crate) use self::r#document::*;\n",
            "pub use crate::numbers::{document::NumbersDocument};\n",
            "pub use {document::*};\n",
            "pub use {self::sheet::*};\n",
            "pub(crate) mod sheet;\n",
            "pub use self::sheet::NumbersSheet;\n",
        )
        for declaration in declarations:
            with self.subTest(declaration=declaration):
                with tempfile.TemporaryDirectory() as directory:
                    root = Path(directory)
                    module = root / boundaries.IWA_NUMBERS_MODULE_SOURCE
                    module.parent.mkdir(parents=True)
                    module.write_text(declaration, encoding="utf-8")
                    self.assertTrue(
                        boundaries.audit_iwa_numbers_document_source_topology(root)
                    )

    def test_retired_iwa_numbers_document_multiline_aliases_cannot_return(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source = root / "crates/adapter/src/lib.rs"
            source.parent.mkdir(parents=True)
            source.write_text(
                "pub\ntype\nLegacy\n=\nNumbersDocumentStats\n;\n",
                encoding="utf-8",
            )
            violations = boundaries.audit_iwa_numbers_document_source_topology(root)
            joined = "\n".join(violations)
            self.assertIn("workspace public name NumbersDocumentStats", joined)
            self.assertIn("workspace type usage NumbersDocumentStats", joined)

    def test_iwa_numbers_document_reader_policy_allows_builder_editor_and_direct_focused_use(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source = root / "crates/litchi-iwa/src/numbers/creation.rs"
            source.parent.mkdir(parents=True)
            source.write_text(
                "pub struct NumbersDocumentBuilder;\n"
                "pub struct NumbersEditor;\n"
                "pub struct NumbersTable;\n"
                "pub struct TableDataExtractor;\n"
                "fn direct() -> (litchi_numbers::Document, litchi_numbers::Package, "
                "litchi_numbers::Sheet) "
                "{ todo!() }\n",
                encoding="utf-8",
            )
            self.assertEqual(
                boundaries.audit_iwa_numbers_document_source_topology(root), []
            )

    def test_iwa_numbers_document_reader_policy_rejects_host_focused_facades(self) -> None:
        declarations = (
            "pub use litchi_numbers::Document;\n",
            "pub use litchi_numbers::{Package as Spreadsheet};\n",
            "pub use litchi_numbers::*;\n",
            "pub use litchi_numbers as numbers_api;\n",
            "pub use litchi_numbers::document as reader;\n",
            "pub use litchi_numbers::{package as artifact};\n",
            "use litchi_numbers::Document as Focused;\npub type Reader = Focused;\n",
            "use litchi_numbers::document as semantic;\npub type Reader = semantic::Document;\n",
            "pub type Reader = litchi_numbers::Document;\n",
        )
        for declaration in declarations:
            with self.subTest(declaration=declaration):
                with tempfile.TemporaryDirectory() as directory:
                    root = Path(directory)
                    source = root / "crates/litchi-iwa/src/numbers/facade.rs"
                    source.parent.mkdir(parents=True)
                    source.write_text(declaration, encoding="utf-8")
                    violations = boundaries.audit_iwa_numbers_document_source_topology(root)
                    self.assertTrue(violations)
                    self.assertTrue(
                        all("focused host facade" in item for item in violations),
                        violations,
                    )

    def test_retired_numbers_document_public_names_are_workspace_wide(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source = root / "crates/unrelated/src/lib.rs"
            source.parent.mkdir(parents=True)
            source.write_text(
                "pub struct NumbersDocument;\n"
                "fn private_reader(_: NumbersSheet) {}\n"
                "/// [`NumbersDocumentStats`] must stay retired.\n"
                "pub struct Report;\n",
                encoding="utf-8",
            )
            violations = boundaries.audit_iwa_numbers_document_source_topology(root)
            joined = "\n".join(violations)
            self.assertIn("workspace public name NumbersDocument", joined)
            self.assertIn("workspace type usage NumbersSheet", joined)
            self.assertIn("workspace public rustdoc NumbersDocumentStats", joined)

    def test_focused_numbers_document_reader_public_api_rejects_native_leaks(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            document = root / boundaries.NUMBERS_DOCUMENT_PUBLIC_API_SOURCES[0]
            document.parent.mkdir(parents=True)
            document.write_text(
                "/// Returns a `NativeObjectIdentifier`.\n"
                "pub struct Document { pub object_id: u64 }\n"
                "pub fn raw_source_bytes() -> &[u8] { todo!() }\n"
                "pub type Native = litchi_iwa_common::ObjectArchive;\n",
                encoding="utf-8",
            )
            violations = boundaries.audit_numbers_document_public_api(root)
            joined = "\n".join(violations)
            self.assertIn("raw identifier object_id", joined)
            self.assertIn("raw source bytes raw_source_bytes", joined)
            self.assertIn("archive/IWA type litchi_iwa_common", joined)
            self.assertIn("native object ObjectArchive", joined)
            self.assertIn("rustdoc exposes raw identifier NativeObjectIdentifier", joined)

    def test_focused_numbers_document_reader_public_api_allows_checked_limits(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            document = root / boundaries.NUMBERS_DOCUMENT_PUBLIC_API_SOURCES[0]
            document.parent.mkdir(parents=True)
            document.write_text(
                "/// Immutable semantic spreadsheet.\n"
                "pub struct Document;\n"
                "pub struct DocumentReadOptions;\n"
                "pub struct DocumentSourceLimits;\n"
                "pub enum DocumentSourceLimitKind { InputBytes }\n"
                "pub enum ReadLimitKind { Sheets, Tables, Cells, TextBytes }\n"
                "pub struct Stats { pub sheet_count: usize }\n"
                "pub fn from_shared_bytes(_: std::sync::Arc<[u8]>) -> Document "
                "{ Document }\n",
                encoding="utf-8",
            )
            self.assertEqual(boundaries.audit_numbers_document_public_api(root), [])

    def test_retired_iwa_pages_document_reader_inventory_is_exact(self) -> None:
        self.assertEqual(
            boundaries.RETIRED_IWA_PAGES_DOCUMENT_SOURCE,
            Path("crates/litchi-iwa/src/pages/document.rs"),
        )
        self.assertEqual(
            boundaries.RETIRED_IWA_PAGES_DOCUMENT_TYPES,
            ("PagesDocument", "PagesDocumentState", "PagesDocumentStats"),
        )
        self.assertEqual(
            boundaries.IWA_PAGES_MODULE_SOURCE,
            Path("crates/litchi-iwa/src/pages/mod.rs"),
        )
        self.assertEqual(boundaries.WORKSPACE_CRATES_ROOT, Path("crates"))
        self.assertEqual(
            boundaries.IWA_PAGES_FOCUSED_READER_TYPES,
            frozenset({"Document", "Package"}),
        )

    def test_retired_iwa_pages_document_reader_surface_cannot_return(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            retired = root / boundaries.RETIRED_IWA_PAGES_DOCUMENT_SOURCE
            retired.parent.mkdir(parents=True)
            retired.write_text("// retired reader returned\n", encoding="utf-8")
            module = root / boundaries.IWA_PAGES_MODULE_SOURCE
            module.write_text(
                "pub mod r#document;\n"
                "pub use self::r#document::*;\n",
                encoding="utf-8",
            )
            source_caller = root / "crates/litchi-iwa/src/legacy_pages.rs"
            source_caller.parent.mkdir(parents=True, exist_ok=True)
            source_caller.write_text(
                "pub fn open() -> PagesDocument { todo!() }\n"
                "pub type State = PagesDocumentState;\n"
                "/// Do not restore `PagesDocumentStats`.\n",
                encoding="utf-8",
            )
            test_caller = root / "crates/litchi-iwa/tests/pages_reader.rs"
            test_caller.parent.mkdir(parents=True)
            test_caller.write_text(
                "fn assert_stats(_: PagesDocumentStats) {}\n",
                encoding="utf-8",
            )
            example_caller = root / "crates/litchi-iwa/examples/read_pages.rs"
            example_caller.parent.mkdir(parents=True)
            example_caller.write_text(
                "use litchi_iwa::pages::PagesDocument;\n",
                encoding="utf-8",
            )
            readme = root / boundaries.IWA_PAGES_README
            readme.write_text(
                "Open with `PagesDocumentState` and inspect "
                "`PagesDocumentStats`.\n",
                encoding="utf-8",
            )

            violations = boundaries.audit_iwa_pages_document_source_topology(root)
            expected = {
                        "retired litchi-iwa Pages document reader local re-export "
                        "document: crates/litchi-iwa/src/pages/mod.rs:2",
                        "retired litchi-iwa Pages document reader module document: "
                        "crates/litchi-iwa/src/pages/mod.rs:1",
                        "retired litchi-iwa Pages document reader README reference "
                        "PagesDocumentState: crates/litchi-iwa/README.md:1",
                        "retired litchi-iwa Pages document reader README reference "
                        "PagesDocumentStats: crates/litchi-iwa/README.md:1",
                        "retired litchi-iwa Pages document reader source returned: "
                        "crates/litchi-iwa/src/pages/document.rs",
                        "retired litchi-iwa Pages document reader rustdoc reference "
                        "PagesDocumentStats: crates/litchi-iwa/src/legacy_pages.rs:3",
                        "retired litchi-iwa Pages document reader type usage "
                        "PagesDocument: crates/litchi-iwa/examples/read_pages.rs:1",
                        "retired litchi-iwa Pages document reader type usage "
                        "PagesDocument: crates/litchi-iwa/src/legacy_pages.rs:1",
                        "retired litchi-iwa Pages document reader type usage "
                        "PagesDocumentState: crates/litchi-iwa/src/legacy_pages.rs:2",
                        "retired litchi-iwa Pages document reader type usage "
                        "PagesDocumentStats: crates/litchi-iwa/tests/pages_reader.rs:1",
            }
            self.assertTrue(expected <= set(violations), violations)
            self.assertTrue(
                any(
                    "workspace public name PagesDocument" in violation
                    for violation in violations
                ),
                violations,
            )

    def test_retired_iwa_pages_document_rustdoc_references_cannot_return(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            path = root / boundaries.IWA_PAGES_MODULE_SOURCE
            path.parent.mkdir(parents=True)
            path.write_text(
                "//! Use `PagesDocument`.\n"
                "/// State was `PagesDocumentState`.\n"
                "/** Stats were `PagesDocumentStats`. */\n"
                '#[doc = "Do not restore PagesDocument"]\n'
                "pub struct Reader;\n",
                encoding="utf-8",
            )

            violations = boundaries.audit_iwa_pages_document_source_topology(root)
            expected = {
                    "retired litchi-iwa Pages document reader rustdoc reference "
                    "PagesDocument: crates/litchi-iwa/src/pages/mod.rs:1",
                    "retired litchi-iwa Pages document reader rustdoc reference "
                    "PagesDocument: crates/litchi-iwa/src/pages/mod.rs:4",
                    "retired litchi-iwa Pages document reader rustdoc reference "
                    "PagesDocumentState: crates/litchi-iwa/src/pages/mod.rs:2",
                    "retired litchi-iwa Pages document reader rustdoc reference "
                    "PagesDocumentStats: crates/litchi-iwa/src/pages/mod.rs:3",
            }
            self.assertTrue(expected <= set(violations), violations)
            self.assertEqual(
                sum("workspace public rustdoc" in item for item in violations),
                4,
            )

    def test_retired_iwa_pages_document_module_and_reexport_variants(
        self,
    ) -> None:
        declarations = (
            ("mod document;", "module document"),
            ("pub(crate) mod r#document;", "module document"),
            ("pub\nmod\ndocument\n{}", "module document"),
            ("pub use document::*;", "local re-export document"),
            ("pub(crate) use self::r#document as legacy;", "local re-export document"),
            (
                "pub\nuse\ncrate::r#pages::document\n    as LegacyReader;",
                "local re-export document",
            ),
            (
                "pub use crate::pages::{r#document as LegacyReader};",
                "local re-export document",
            ),
        )
        for declaration, fragment in declarations:
            with self.subTest(declaration=declaration):
                with tempfile.TemporaryDirectory() as directory:
                    root = Path(directory)
                    module = root / boundaries.IWA_PAGES_MODULE_SOURCE
                    module.parent.mkdir(parents=True)
                    module.write_text(declaration + "\n", encoding="utf-8")

                    violations = (
                        boundaries.audit_iwa_pages_document_source_topology(root)
                    )
                    self.assertTrue(violations)
                    self.assertTrue(
                        any(fragment in violation for violation in violations),
                        violations,
                    )

    def test_retired_iwa_pages_document_multiline_type_aliases_cannot_return(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            caller = root / "crates/litchi-iwa/src/legacy_pages_aliases.rs"
            caller.parent.mkdir(parents=True)
            caller.write_text(
                "use litchi_iwa\n"
                "    ::pages\n"
                "    ::PagesDocument\n"
                "    as LegacyReader;\n"
                "pub type LegacyState =\n"
                "    PagesDocumentState;\n"
                "fn stats() -> crate::pages\n"
                "    ::PagesDocumentStats { todo!() }\n",
                encoding="utf-8",
            )

            violations = boundaries.audit_iwa_pages_document_source_topology(root)
            expected = {
                    "retired litchi-iwa Pages document reader type usage "
                    "PagesDocument: "
                    "crates/litchi-iwa/src/legacy_pages_aliases.rs:3",
                    "retired litchi-iwa Pages document reader type usage "
                    "PagesDocumentState: "
                    "crates/litchi-iwa/src/legacy_pages_aliases.rs:6",
                    "retired litchi-iwa Pages document reader type usage "
                    "PagesDocumentStats: "
                    "crates/litchi-iwa/src/legacy_pages_aliases.rs:8",
            }
            self.assertTrue(expected <= set(violations), violations)
            self.assertTrue(
                any(
                    "workspace public name PagesDocumentState" in item
                    for item in violations
                ),
                violations,
            )

    def test_iwa_pages_document_reader_policy_allows_builder_editor_and_direct_focused_use(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            module = root / boundaries.IWA_PAGES_MODULE_SOURCE
            module.parent.mkdir(parents=True)
            module.write_text(
                "pub use creation::PagesDocumentBuilder;\n"
                "pub use editor::PagesEditor;\n"
                "use litchi_pages::{Document as FocusedDocument, "
                "Package as FocusedPackage};\n"
                "fn read(_: FocusedDocument, _: FocusedPackage) {}\n",
                encoding="utf-8",
            )
            for relative in (
                Path("crates/litchi-iwa/src/pages/creation.rs"),
                Path("crates/litchi-iwa/tests/generated_roundtrip.rs"),
                Path("crates/litchi-iwa/examples/create_pages.rs"),
            ):
                path = root / relative
                path.parent.mkdir(parents=True, exist_ok=True)
                path.write_text(
                    "use litchi_iwa::pages::{PagesDocumentBuilder, PagesEditor};\n"
                    "use litchi_pages::{Document, Package};\n"
                    "pub struct PagesDocuments;\n"
                    "pub struct LegacyPagesDocumentStats;\n"
                    "// PagesDocument and PagesDocumentState are retired.\n",
                    encoding="utf-8",
                )
            readme = root / boundaries.IWA_PAGES_README
            readme.write_text(
                "Use `PagesDocumentBuilder`, `PagesEditor`, "
                "`litchi_pages::Document`, or `litchi_pages::Package`.\n",
                encoding="utf-8",
            )

            self.assertEqual(
                boundaries.audit_iwa_pages_document_source_topology(root), []
            )

    def test_iwa_pages_document_reader_policy_rejects_host_focused_facades(
        self,
    ) -> None:
        declarations = (
            ("pub use litchi_pages::Document;", "facade Document"),
            ("pub use litchi_pages::Package as PagesReader;", "facade PagesReader"),
            (
                "use litchi_pages::Document as FocusedDocument;\n"
                "pub type Reader = FocusedDocument;",
                "facade FocusedDocument",
            ),
            (
                "pub fn open() -> litchi_pages::Document { todo!() }",
                "facade Document",
            ),
            (
                "use litchi_pages as focused;\n"
                "pub type Reader = focused::Document;",
                "facade Document",
            ),
            (
                "use litchi_pages::Document as Focused;\n"
                "type Inner = Focused;\n"
                "pub fn read() -> Inner { todo!() }",
                "facade Inner",
            ),
            (
                "pub use litchi_pages as focused_pages;",
                "facade focused_pages",
            ),
            (
                "pub use litchi_pages::{self as focused_pages};",
                "facade focused_pages",
            ),
            ("pub use litchi_pages::*;", "facade glob"),
        )
        for declaration, fragment in declarations:
            with self.subTest(declaration=declaration):
                with tempfile.TemporaryDirectory() as directory:
                    root = Path(directory)
                    host = root / boundaries.IWA_PAGES_MODULE_SOURCE
                    host.parent.mkdir(parents=True)
                    host.write_text(declaration + "\n", encoding="utf-8")

                    violations = (
                        boundaries.audit_iwa_pages_document_source_topology(root)
                    )
                    self.assertTrue(violations)
                    self.assertTrue(
                        any(fragment in violation for violation in violations),
                        violations,
                    )

    def test_retired_pages_document_public_names_are_workspace_wide(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            aliases = root / "crates/consumer/src/lib.rs"
            aliases.parent.mkdir(parents=True)
            aliases.write_text(
                "pub use litchi_pages::Document as PagesDocument;\n"
                "pub type PagesDocumentStats = crate::Stats;\n"
                "pub fn reopen(document: PagesDocument) {}\n"
                "/// The old PagesDocumentState must stay gone.\n"
                "pub struct Reader;\n",
                encoding="utf-8",
            )
            private_use = root / "crates/other/src/lib.rs"
            private_use.parent.mkdir(parents=True)
            private_use.write_text(
                "use litchi_pages::Document as PagesDocument;\n",
                encoding="utf-8",
            )

            self.assertEqual(
                boundaries.audit_iwa_pages_document_source_topology(root),
                [
                    "retired Pages document reader workspace public name "
                    "PagesDocument: crates/consumer/src/lib.rs:1",
                    "retired Pages document reader workspace public name "
                    "PagesDocument: crates/consumer/src/lib.rs:3",
                    "retired Pages document reader workspace public name "
                    "PagesDocumentStats: crates/consumer/src/lib.rs:2",
                    "retired Pages document reader workspace public rustdoc "
                    "PagesDocumentState: crates/consumer/src/lib.rs:4",
                ],
            )

    def test_focused_pages_document_reader_public_api_rejects_native_leaks(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            semantic = root / boundaries.PAGES_DOCUMENT_PUBLIC_API_SOURCES[0]
            semantic.parent.mkdir(parents=True)
            semantic.write_text(
                "/// Returns `NativeObjectId` and `DocumentArchive`.\n"
                "pub fn inspect(object_id: u64) -> DocumentArchive { todo!() }\n"
                "pub type SourceBytes = Vec<u8>;\n",
                encoding="utf-8",
            )
            export = root / boundaries.PAGES_DOCUMENT_PUBLIC_API_SOURCES[1]
            export.write_text(
                "pub type DocumentReadOptions = litchi_iwa_core::RawObject;\n",
                encoding="utf-8",
            )

            violations = boundaries.audit_pages_document_public_api(root)
            self.assertEqual(violations, sorted(violations))
            self.assertTrue(
                any("raw identifier object_id" in violation for violation in violations)
            )
            self.assertTrue(
                any("archive/IWA type DocumentArchive" in violation for violation in violations)
            )
            self.assertTrue(
                any("raw source bytes SourceBytes" in violation for violation in violations)
            )
            self.assertTrue(
                any("archive/IWA type litchi_iwa_core" in violation for violation in violations)
            )
            self.assertTrue(
                any("native object RawObject" in violation for violation in violations)
            )
            self.assertTrue(
                any("rustdoc exposes raw identifier NativeObjectId" in violation for violation in violations)
            )

    def test_focused_pages_document_reader_public_api_allows_checked_limits(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            semantic = root / boundaries.PAGES_DOCUMENT_PUBLIC_API_SOURCES[0]
            semantic.parent.mkdir(parents=True)
            semantic.write_text(
                "/// Checked, content-free source limits.\n"
                "pub struct DocumentSourceLimits { max_input_bytes: u64 }\n"
                "pub struct DocumentReadOptions;\n"
                "pub enum ReadError { InvalidSource }\n"
                "pub struct Document;\n",
                encoding="utf-8",
            )
            export = root / boundaries.PAGES_DOCUMENT_PUBLIC_API_SOURCES[1]
            export.write_text(
                "pub use document::{Document, DocumentReadOptions, "
                "DocumentSourceLimits, ReadError};\n",
                encoding="utf-8",
            )

            self.assertEqual(boundaries.audit_pages_document_public_api(root), [])

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


    def test_numbers_table_title_settings_boundary_inventories_are_exact(self) -> None:
        self.assertEqual(
            boundaries.RETIRED_IWA_NUMBERS_TABLE_TITLE_SETTINGS_METHODS,
            ("table_title_settings", "set_table_title_settings"),
        )
        self.assertEqual(
            boundaries.RETIRED_IWA_NUMBERS_TABLE_TITLE_SETTINGS_EXAMPLE,
            Path("crates/litchi-iwa/examples/edit_numbers_table_title.rs"),
        )
        self.assertEqual(
            boundaries.RETIRED_IWA_NUMBERS_TABLE_TITLE_SETTINGS_TESTS,
            (
                "table_title_settings_are_lossless_transactional_and_wire_exact",
                "table_title_settings_restore_native_presence_exactly",
                "table_title_settings_reject_missing_render_styles_transactionally",
                "table_title_settings_reject_malformed_wire_transactionally",
            ),
        )
        self.assertEqual(
            boundaries.NUMBERS_TABLE_TITLE_SETTINGS_SEMANTIC_SOURCE,
            Path("crates/litchi-numbers/src/table/title.rs"),
        )
        self.assertEqual(
            boundaries.NUMBERS_TABLE_TITLE_SETTINGS_OWNER_SOURCE,
            Path("crates/litchi-numbers/src/package/table_title.rs"),
        )
        self.assertEqual(
            boundaries.NUMBERS_TABLE_TITLE_SETTINGS_OWNER_HELPER_ROOT,
            Path("crates/litchi-numbers/src/package/table_title"),
        )
        self.assertEqual(
            boundaries.NUMBERS_TABLE_TITLE_SETTINGS_CANONICAL_TYPES,
            (
                "Settings",
                "Edit",
                "Patch",
                "Commit",
                "Diagnostics",
                "Error",
                "LimitKind",
                "Path",
            ),
        )
        self.assertEqual(
            boundaries.NUMBERS_TABLE_TITLE_SETTINGS_PACKAGE_METHODS,
            ("table_title_settings", "edit_table_title", "apply_table_title"),
        )
        self.assertTrue(
            {
                "TitleSettings",
                "TableTitleSettings",
                "TableTitleEdit",
                "TableTitlePatch",
                "TableTitleCommit",
            }
            <= boundaries.NUMBERS_TABLE_TITLE_SETTINGS_FLAT_ALIASES
        )

    def test_retired_iwa_numbers_table_title_surface_cannot_return(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            host = root / boundaries.IWA_NUMBERS_SOURCE_ROOT / "editor/table_title.rs"
            host.parent.mkdir(parents=True)
            host.write_text(
                "pub fn table_title_settings() {}\n"
                "pub fn set_table_title_settings() {}\n"
                "pub(crate) fn table_title_settings_in_package() {}\n"
                "pub(crate) fn set_table_title_settings_in_package() {}\n",
                encoding="utf-8",
            )
            tests = root / boundaries.IWA_NUMBERS_EDITOR_TEST_SOURCE
            tests.write_text(
                "\n".join(
                    f"fn {name}() {{}}"
                    for name in (
                        boundaries.RETIRED_IWA_NUMBERS_TABLE_TITLE_SETTINGS_TESTS
                    )
                )
                + "\n",
                encoding="utf-8",
            )
            example = root / boundaries.RETIRED_IWA_NUMBERS_TABLE_TITLE_SETTINGS_EXAMPLE
            example.parent.mkdir(parents=True)
            example.write_text("fn main() {}\n", encoding="utf-8")

            violations = (
                boundaries.audit_iwa_numbers_table_title_settings_source_topology(
                    root
                )
            )

            self.assertEqual(len(violations), 7)
            self.assertTrue(any("example returned" in item for item in violations))
            for method in boundaries.RETIRED_IWA_NUMBERS_TABLE_TITLE_SETTINGS_METHODS:
                self.assertTrue(
                    any(f"settings method {method}:" in item for item in violations)
                )
            for name in boundaries.RETIRED_IWA_NUMBERS_TABLE_TITLE_SETTINGS_TESTS:
                self.assertTrue(
                    any(f"settings test {name}:" in item for item in violations)
                )
            self.assertFalse(any("in_package" in item for item in violations))

    def test_retired_iwa_numbers_table_title_readme_calls_and_example(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            readme = root / boundaries.IWA_NUMBERS_README
            readme.parent.mkdir(parents=True)
            readme.write_text(
                "numbers.table_title_settings(sheet, table);\n"
                "numbers_editor\n  .\n  "
                "set_table_title_settings(sheet, table, value);\n"
                "crate::NumbersEditor::table_title_settings(sheet, table);\n"
                "edit_numbers_table_title\n"
                "edit_numbers_table_title.rs\n",
                encoding="utf-8",
            )

            violations = (
                boundaries.audit_iwa_numbers_table_title_settings_source_topology(
                    root
                )
            )

            self.assertEqual(len(violations), 5)
            self.assertEqual(
                sum("README call" in item for item in violations),
                3,
            )
            self.assertEqual(
                sum("README example reference" in item for item in violations),
                2,
            )

    def test_iwa_numbers_table_title_policy_retains_private_shared_helpers(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            host = root / boundaries.IWA_NUMBERS_SOURCE_ROOT / "editor/table_title.rs"
            host.parent.mkdir(parents=True)
            host.write_text(
                "pub(crate) fn table_title_settings_in_package() {}\n"
                "pub(crate) fn set_table_title_settings_in_package() {}\n"
                "pub(super) fn read_table_title_settings_wire() {}\n"
                "pub(super) fn write_table_title_settings_wire() {}\n",
                encoding="utf-8",
            )
            pages = root / "crates/litchi-iwa/src/pages/editor/table_title.rs"
            pages.parent.mkdir(parents=True)
            pages.write_text(
                "use crate::numbers::editor::{table_title_settings_in_package, "
                "set_table_title_settings_in_package};\n",
                encoding="utf-8",
            )
            readme = root / boundaries.IWA_NUMBERS_README
            readme.parent.mkdir(parents=True, exist_ok=True)
            readme.write_text(
                "Use table::title and package.edit_table_title().\n",
                encoding="utf-8",
            )

            self.assertEqual(
                boundaries.audit_iwa_numbers_table_title_settings_source_topology(root),
                [],
            )

    def test_focused_numbers_table_title_requires_each_direct_canonical_type(
        self,
    ) -> None:
        for missing in boundaries.NUMBERS_TABLE_TITLE_SETTINGS_CANONICAL_TYPES:
            with self.subTest(missing=missing):
                with tempfile.TemporaryDirectory() as directory:
                    root = Path(directory)
                    add_numbers_table_title_settings_canonical_scaffold(root)
                    semantic = (
                        root
                        / boundaries.NUMBERS_TABLE_TITLE_SETTINGS_SEMANTIC_SOURCE
                    )
                    semantic.write_text(
                        "".join(
                            f"pub struct {name};\n"
                            for name in (
                                boundaries.NUMBERS_TABLE_TITLE_SETTINGS_CANONICAL_TYPES
                            )
                            if name != missing
                        ),
                        encoding="utf-8",
                    )

                    self.assertEqual(
                        boundaries.audit_numbers_table_title_settings_facade_source_topology(
                            root
                        ),
                        [
                            "focused litchi-numbers table-title settings public API is "
                            f"missing canonical table::title type {missing}: "
                            "crates/litchi-numbers/src/table/title.rs"
                        ],
                    )

    def test_focused_numbers_table_title_requires_modules_and_private_owner(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            add_numbers_table_title_settings_canonical_scaffold(root)
            package = root / boundaries.NUMBERS_TABLE_TITLE_SETTINGS_EXPORT_SOURCES[1]
            package.write_text("pub mod table_title;\n", encoding="utf-8")
            table = root / boundaries.NUMBERS_TABLE_TITLE_SETTINGS_EXPORT_SOURCES[2]
            table.write_text("mod title;\n", encoding="utf-8")

            violations = (
                boundaries.audit_numbers_table_title_settings_facade_source_topology(
                    root
                )
            )

            self.assertTrue(
                any(
                    "missing canonical table::title module" in item
                    for item in violations
                )
            )
            self.assertTrue(
                any(
                    "exposes duplicate package::table_title module" in item
                    for item in violations
                )
            )

        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            add_numbers_table_title_settings_canonical_scaffold(root)
            (root / boundaries.NUMBERS_TABLE_TITLE_SETTINGS_OWNER_SOURCE).unlink()
            violations = (
                boundaries.audit_numbers_table_title_settings_facade_source_topology(
                    root
                )
            )
            self.assertEqual(
                violations,
                [
                    "focused litchi-numbers table-title settings public API is "
                    "missing private package owner source: "
                    "crates/litchi-numbers/src/package/table_title.rs"
                ],
            )

    def test_focused_numbers_table_title_rejects_aliases_and_physical_leaks(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            add_numbers_table_title_settings_canonical_scaffold(root)
            helper = (
                root
                / boundaries.NUMBERS_TABLE_TITLE_SETTINGS_OWNER_HELPER_ROOT
                / "api.rs"
            )
            helper.parent.mkdir(parents=True)
            helper.write_text(
                "pub type TableTitleEdit = Edit;\n"
                "pub fn table_title_settings(object_id: u64, source_bytes: &[u8], "
                "wire: WireView, archive: Archive, generated: GeneratedProjection, "
                "buffa: BuffaView, prost: prost_types::MessageInfo) {}\n",
                encoding="utf-8",
            )
            lib = root / boundaries.NUMBERS_TABLE_TITLE_SETTINGS_EXPORT_SOURCES[0]
            lib.write_text(
                "pub mod table;\n"
                "pub use crate::table::title::{Settings, Edit};\n",
                encoding="utf-8",
            )

            violations = (
                boundaries.audit_numbers_table_title_settings_facade_source_topology(
                    root
                )
            )

            for fragment in (
                "retains flat alias TableTitleEdit",
                "exposes raw identifier object_id",
                "exposes raw source bytes source_bytes",
                "exposes raw byte slice &[u8]",
                "exposes wire type WireView",
                "exposes archive/IWA type Archive",
                "exposes generated type GeneratedProjection",
                "exposes protobuf type BuffaView",
                "exposes protobuf type prost",
                "exposes protobuf type prost_types",
                "exposes public table-title owner alias",
                "retains root alias Settings",
                "retains root alias Edit",
            ):
                self.assertTrue(
                    any(fragment in item for item in violations),
                    msg=f"missing violation containing {fragment!r}: {violations!r}",
                )

    def test_focused_numbers_table_title_allows_canonical_common_settings_reexport(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            add_numbers_table_title_settings_canonical_scaffold(root)
            semantic = root / boundaries.NUMBERS_TABLE_TITLE_SETTINGS_SEMANTIC_SOURCE
            semantic.write_text(
                "pub use crate::package::table_title::{Commit, Diagnostics, "
                "Edit, Error, "
                "LimitKind, Patch, Path};\n"
                "pub use litchi_iwa_common::table::title::Settings;\n",
                encoding="utf-8",
            )

            self.assertEqual(
                boundaries.audit_numbers_table_title_settings_facade_source_topology(
                    root
                ),
                [],
            )

    def test_numbers_table_dimension_boundary_inventories_are_exact(self) -> None:
        self.assertEqual(
            boundaries.RETIRED_IWA_NUMBERS_TABLE_DIMENSION_METHODS,
            (
                "table_dimension_size",
                "set_table_dimension_size",
                "table_row_height",
                "set_table_row_height",
                "table_column_width",
                "set_table_column_width",
            ),
        )
        self.assertEqual(
            boundaries.RETIRED_IWA_NUMBERS_TABLE_DIMENSION_TYPES,
            ("Dimension", "Points", "Size"),
        )
        self.assertEqual(
            boundaries.RETIRED_IWA_NUMBERS_TABLE_DIMENSION_EXAMPLE,
            Path("crates/litchi-iwa/examples/edit_numbers_table_dimension.rs"),
        )
        self.assertEqual(
            boundaries.RETIRED_IWA_NUMBERS_TABLE_DIMENSION_TESTS,
            (
                "table_dimension_sizes_are_typed_transactional_and_wire_exact",
                "table_dimension_size_preserves_unknown_header_fields",
                "table_dimension_size_rejects_malformed_headers_transactionally",
            ),
        )
        self.assertEqual(
            boundaries.NUMBERS_TABLE_DIMENSION_TRANSACTION_SOURCE,
            Path("crates/litchi-numbers/src/table/dimension/transaction.rs"),
        )
        self.assertEqual(
            boundaries.NUMBERS_TABLE_DIMENSION_OWNER_SOURCE,
            Path("crates/litchi-numbers/src/package/table_dimension.rs"),
        )
        self.assertEqual(
            boundaries.NUMBERS_TABLE_DIMENSION_TRANSACTION_TYPES,
            (
                "Edit",
                "Patch",
                "Commit",
                "Diagnostics",
                "Path",
                "LimitKind",
                "TransactionError",
            ),
        )
        self.assertEqual(
            boundaries.NUMBERS_TABLE_DIMENSION_PACKAGE_METHODS,
            (
                "table_dimension_size",
                "edit_table_dimension_size",
                "apply_table_dimension_size",
            ),
        )

    def test_retired_iwa_numbers_table_dimension_public_surface_cannot_return(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            host = root / boundaries.IWA_NUMBERS_SOURCE_ROOT / "editor/table_dimension.rs"
            host.parent.mkdir(parents=True)
            host.write_text(
                "impl NumbersEditor {\n"
                + "".join(
                    f"pub fn {name}(&self, value: Dimension) -> Size {{ todo!() }}\n"
                    for name in boundaries.RETIRED_IWA_NUMBERS_TABLE_DIMENSION_METHODS
                )
                + "}\n"
                "pub(super) fn set_attached_table_dimension_size() {}\n"
                "pub(super) fn read_attached_table_dimension_size() {}\n",
                encoding="utf-8",
            )
            editor_facade = root / boundaries.IWA_NUMBERS_SOURCE_ROOT / "editor.rs"
            editor_facade.write_text(
                "pub use litchi_numbers::table::dimension::{\n"
                "Dimension, Points, Size\n"
                "};\n",
                encoding="utf-8",
            )
            root_facade = root / boundaries.IWA_NUMBERS_SOURCE_ROOT / "mod.rs"
            root_facade.write_text(
                "pub use editor::{Dimension, Points, Size};\n",
                encoding="utf-8",
            )
            example = root / boundaries.RETIRED_IWA_NUMBERS_TABLE_DIMENSION_EXAMPLE
            example.parent.mkdir(parents=True)
            example.write_text("fn main() {}\n", encoding="utf-8")
            tests = root / boundaries.IWA_NUMBERS_EDITOR_TEST_SOURCE
            tests.write_text(
                "".join(
                    f"fn {name}() {{}}\n"
                    for name in boundaries.RETIRED_IWA_NUMBERS_TABLE_DIMENSION_TESTS
                )
                + "fn table_dimension_private_helpers_stay_covered() {}\n",
                encoding="utf-8",
            )

            violations = boundaries.audit_iwa_numbers_table_dimension_source_topology(
                root
            )

            self.assertTrue(any("example returned" in item for item in violations))
            for method in boundaries.RETIRED_IWA_NUMBERS_TABLE_DIMENSION_METHODS:
                self.assertTrue(
                    any(f"public table-dimension method {method}:" in item for item in violations),
                    msg=f"missing retired method {method!r}: {violations!r}",
                )
            for exposed in ("Dimension", "Points", "Size"):
                self.assertTrue(
                    any(f"public facade {exposed}:" in item for item in violations),
                    msg=f"missing host facade {exposed!r}: {violations!r}",
                )
            for name in boundaries.RETIRED_IWA_NUMBERS_TABLE_DIMENSION_TESTS:
                self.assertTrue(
                    any(f"table-dimension test {name}:" in item for item in violations),
                    msg=f"missing retired test {name!r}: {violations!r}",
                )
            self.assertFalse(any("attached_table" in item for item in violations))
            self.assertFalse(any("private_helpers" in item for item in violations))

    def test_iwa_numbers_table_dimension_rejects_laundered_focused_facades(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            host = root / boundaries.IWA_NUMBERS_SOURCE_ROOT / "facade.rs"
            host.parent.mkdir(parents=True)
            host.write_text(
                "use litchi_numbers::table::dimension::{Dimension as Axis};\n"
                "use litchi_numbers::table::dimension as sizing;\n"
                "type HiddenAxis = Axis;\n"
                "pub type PublicAxis = HiddenAxis;\n"
                "pub use sizing as public_sizing;\n"
                "pub use litchi_numbers as focused_numbers;\n",
                encoding="utf-8",
            )

            violations = boundaries.audit_iwa_numbers_table_dimension_source_topology(
                root
            )

            for exposed in ("HiddenAxis", "sizing", "focused_numbers"):
                self.assertTrue(
                    any(f"public facade {exposed}:" in item for item in violations),
                    msg=f"missing laundered facade {exposed!r}: {violations!r}",
                )

    def test_iwa_numbers_table_dimension_allows_private_shared_physical_helpers(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            host = root / boundaries.IWA_NUMBERS_SOURCE_ROOT / "mod.rs"
            host.parent.mkdir(parents=True)
            host.write_text(
                "pub(crate) use litchi_numbers::table::dimension::{\n"
                "Dimension as NumbersTableDimension,\n"
                "Points as NumbersTablePoints,\n"
                "Size as NumbersTableDimensionSize,\n"
                "};\n"
                "pub(super) fn set_attached_table_dimension_size() {}\n"
                "pub(super) fn read_attached_table_dimension_size() {}\n"
                "pub(crate) fn table_dimension_size_in_package() {}\n"
                "pub(crate) fn set_table_dimension_size_in_package() {}\n",
                encoding="utf-8",
            )
            canonical = root / "crates/litchi/src/numbers.rs"
            canonical.parent.mkdir(parents=True)
            canonical.write_text(
                "pub use litchi_numbers::{Dimension, Points, Size};\n",
                encoding="utf-8",
            )

            self.assertEqual(
                boundaries.audit_iwa_numbers_table_dimension_source_topology(root),
                [],
            )

    def test_focused_numbers_table_dimension_requires_canonical_contract(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            add_numbers_table_dimension_canonical_scaffold(root)

            self.assertEqual(
                boundaries.audit_numbers_table_dimension_facade_source_topology(root),
                [],
            )

            transaction = root / boundaries.NUMBERS_TABLE_DIMENSION_TRANSACTION_SOURCE
            transaction.write_text("pub struct Edit;\n", encoding="utf-8")
            owner = root / boundaries.NUMBERS_TABLE_DIMENSION_OWNER_SOURCE
            owner.write_text(
                "impl Package { pub fn table_dimension_size(&self) {} }\n",
                encoding="utf-8",
            )
            semantic = root / boundaries.NUMBERS_TABLE_DIMENSION_SEMANTIC_SOURCE
            semantic.write_text(
                "pub struct Dimension;\npub struct Points;\npub struct Size;\n",
                encoding="utf-8",
            )

            violations = boundaries.audit_numbers_table_dimension_facade_source_topology(
                root
            )
            for missing in boundaries.NUMBERS_TABLE_DIMENSION_TRANSACTION_TYPES[1:]:
                self.assertTrue(
                    any(f"transaction type {missing}:" in item for item in violations),
                    msg=f"missing canonical type gate {missing!r}: {violations!r}",
                )
            for missing in boundaries.NUMBERS_TABLE_DIMENSION_PACKAGE_METHODS[1:]:
                self.assertTrue(
                    any(f"Package method {missing}:" in item for item in violations),
                    msg=f"missing Package method gate {missing!r}: {violations!r}",
                )
            self.assertTrue(
                any("missing canonical table::dimension::transaction module" in item for item in violations)
            )

    def test_focused_numbers_table_dimension_rejects_aliases_and_physical_leaks(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            add_numbers_table_dimension_canonical_scaffold(root)
            owner = root / boundaries.NUMBERS_TABLE_DIMENSION_OWNER_SOURCE
            owner.write_text(
                owner.read_text(encoding="utf-8")
                + "pub fn edit_table_dimension_size_raw(\n"
                "object_id: u64, native_object: NativeObjectIdentifier,\n"
                "source_bytes: &[u8], archive: Archive, wire: WireView,\n"
                "generated: GeneratedProjection, prost: prost_types::MessageInfo\n"
                ") {}\n",
                encoding="utf-8",
            )
            transaction = root / boundaries.NUMBERS_TABLE_DIMENSION_TRANSACTION_SOURCE
            transaction.write_text(
                transaction.read_text(encoding="utf-8")
                + "pub type DimensionSizeEdit = Edit;\n",
                encoding="utf-8",
            )
            lib = root / boundaries.NUMBERS_TABLE_DIMENSION_EXPORT_SOURCES[0]
            lib.write_text(
                lib.read_text(encoding="utf-8")
                + "pub use table::dimension::transaction::{Edit, TransactionError};\n",
                encoding="utf-8",
            )

            violations = boundaries.audit_numbers_table_dimension_facade_source_topology(
                root
            )

            for fragment in (
                "retains flat alias DimensionSizeEdit",
                "transaction alias outside table::dimension::transaction Edit",
                "transaction alias outside table::dimension::transaction TransactionError",
                "exposes raw identifier object_id",
                "exposes raw identifier NativeObjectIdentifier",
                "exposes native object native_object",
                "exposes raw source bytes source_bytes",
                "exposes raw byte slice &[u8]",
                "exposes archive/IWA type Archive",
                "exposes wire type WireView",
                "exposes generated type GeneratedProjection",
                "exposes protobuf type prost",
                "exposes protobuf type prost_types",
            ):
                self.assertTrue(
                    any(fragment in item for item in violations),
                    msg=f"missing violation containing {fragment!r}: {violations!r}",
                )

    def test_pages_section_settings_boundary_inventories_are_exact(self) -> None:
        self.assertEqual(
            boundaries.RETIRED_IWA_PAGES_SECTION_SETTINGS_METHODS,
            ("section_settings", "set_section_settings", "set_section_name"),
        )
        self.assertEqual(
            boundaries.RETIRED_IWA_PAGES_SECTION_SETTINGS_TESTS,
            (
                "section_settings_crud_is_lossless_validated_and_transactional",
                "section_settings_reject_zero_starting_page_number_transactionally",
            ),
        )
        self.assertEqual(
            boundaries.PAGES_SECTION_SETTINGS_CANONICAL_TYPES,
            (
                "Edit",
                "Patch",
                "Commit",
                "Diagnostics",
                "Error",
                "LimitKind",
                "Path",
                "DependencyKind",
            ),
        )
        self.assertEqual(
            boundaries.PAGES_SECTION_SETTINGS_PACKAGE_METHODS,
            ("section_settings", "edit_section_settings", "apply_section_settings"),
        )
        self.assertTrue(
            {
                "SectionSettings",
                "SectionSettingsEdit",
                "PagesSectionSettings",
                "PagesSectionSettingsPatch",
            }
            <= boundaries.PAGES_SECTION_SETTINGS_FLAT_ALIASES
        )

    def test_retired_iwa_pages_section_settings_surface_cannot_return(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            host = root / boundaries.IWA_PAGES_SOURCE_ROOT / "editor/section_settings.rs"
            host.parent.mkdir(parents=True)
            host.write_text(
                "pub fn section_settings() {}\n"
                "pub fn set_section_settings() {}\n"
                "pub fn set_section_name() {}\n"
                "pub(super) fn section_background_payload() {}\n",
                encoding="utf-8",
            )
            tests = root / boundaries.IWA_PAGES_EDITOR_TEST_SOURCE
            tests.write_text(
                "\n".join(
                    f"fn {name}() {{}}"
                    for name in boundaries.RETIRED_IWA_PAGES_SECTION_SETTINGS_TESTS
                )
                + "\n",
                encoding="utf-8",
            )
            editor = root / boundaries.IWA_PAGES_EDITOR_SOURCE
            editor.write_text("mod section_settings;\n", encoding="utf-8")
            example = root / boundaries.RETIRED_IWA_PAGES_SECTION_SETTINGS_EXAMPLE
            example.parent.mkdir(parents=True, exist_ok=True)
            example.write_text("fn main() {}\n", encoding="utf-8")

            violations = boundaries.audit_iwa_pages_section_settings_source_topology(
                root
            )

            self.assertEqual(len(violations), 6)
            self.assertTrue(any("example returned" in item for item in violations))
            for method in boundaries.RETIRED_IWA_PAGES_SECTION_SETTINGS_METHODS:
                self.assertTrue(
                    any(f"settings method {method}:" in item for item in violations)
                )
            for name in boundaries.RETIRED_IWA_PAGES_SECTION_SETTINGS_TESTS:
                self.assertTrue(any(f"test {name}:" in item for item in violations))
            self.assertFalse(any("section_background_payload" in item for item in violations))

    def test_retired_iwa_pages_section_settings_readme_tokens(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            readme = root / boundaries.IWA_PAGES_README
            readme.parent.mkdir(parents=True)
            readme.write_text(
                "pages.section_settings(selector);\n"
                "editor\n  .\n  set_section_settings(selector, value);\n"
                "other.set_section_name(selector, name);\n"
                "crate::PagesEditor::section_settings(selector);\n"
                "set_pages_section_settings\n"
                "set_pages_section_settings.rs\n",
                encoding="utf-8",
            )

            violations = boundaries.audit_iwa_pages_section_settings_source_topology(
                root
            )

            self.assertEqual(sum("README call" in item for item in violations), 4)
            self.assertEqual(
                sum("README example reference" in item for item in violations), 2
            )

    def test_iwa_pages_section_settings_policy_retains_adjacent_seams(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            host = root / boundaries.IWA_PAGES_SOURCE_ROOT / "editor/section_wire.rs"
            host.parent.mkdir(parents=True)
            host.write_text(
                "pub(super) fn section_message_data() {}\n"
                "pub(super) fn section_background_payload() {}\n"
                "pub(super) fn set_section_background_payload() {}\n"
                "pub(super) fn validate_section_payload() {}\n"
                "pub fn section_background() {}\n"
                "pub fn set_section_background() {}\n"
                "pub fn section_pagination() {}\n"
                "pub fn edit_section_name() {}\n",
                encoding="utf-8",
            )
            retained = root / boundaries.IWA_PAGES_SOURCE_ROOT / "editor/section_settings.rs"
            retained.write_text(
                "pub(super) fn section_background_payload() {}\n"
                "pub(super) fn set_section_background_payload() {}\n",
                encoding="utf-8",
            )
            editor = root / boundaries.IWA_PAGES_EDITOR_SOURCE
            editor.write_text("mod section_settings;\n", encoding="utf-8")
            readme = root / boundaries.IWA_PAGES_README
            readme.parent.mkdir(parents=True, exist_ok=True)
            readme.write_text(
                "Use litchi_pages::section::settings and edit_section_settings.\n",
                encoding="utf-8",
            )

            self.assertEqual(
                boundaries.audit_iwa_pages_section_settings_source_topology(root), []
            )

    def test_focused_pages_section_settings_requires_each_canonical_type(self) -> None:
        for missing in boundaries.PAGES_SECTION_SETTINGS_CANONICAL_TYPES:
            with self.subTest(missing=missing):
                with tempfile.TemporaryDirectory() as directory:
                    root = Path(directory)
                    add_pages_section_settings_canonical_scaffold(root)
                    semantic = root / boundaries.PAGES_SECTION_SETTINGS_SEMANTIC_SOURCE
                    semantic.write_text(
                        "".join(
                            f"pub struct {name};\n"
                            for name in boundaries.PAGES_SECTION_SETTINGS_CANONICAL_TYPES
                            if name != missing
                        ),
                        encoding="utf-8",
                    )

                    self.assertEqual(
                        boundaries.audit_pages_section_settings_facade_source_topology(
                            root
                        ),
                        [
                            "focused litchi-pages section-settings public API is "
                            f"missing canonical section::settings type {missing}: "
                            "crates/litchi-pages/src/section/settings.rs"
                        ],
                    )

    def test_focused_pages_section_settings_requires_nested_private_owner(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            add_pages_section_settings_canonical_scaffold(root)
            package = root / boundaries.PAGES_SECTION_SETTINGS_EXPORT_SOURCES[1]
            package.write_text("pub mod section_settings;\n", encoding="utf-8")
            section = root / boundaries.PAGES_SECTION_SETTINGS_EXPORT_SOURCES[2]
            section.write_text("mod settings;\n", encoding="utf-8")

            violations = boundaries.audit_pages_section_settings_facade_source_topology(
                root
            )
            self.assertTrue(
                any(
                    "missing canonical section::settings module" in item
                    for item in violations
                )
            )
            self.assertTrue(
                any(
                    "exposes duplicate package::section_settings module" in item
                    for item in violations
                )
            )

        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            add_pages_section_settings_canonical_scaffold(root)
            (root / boundaries.PAGES_SECTION_SETTINGS_OWNER_SOURCE).unlink()
            self.assertEqual(
                boundaries.audit_pages_section_settings_facade_source_topology(root),
                [
                    "focused litchi-pages section-settings public API is missing "
                    "Package method apply_section_settings: "
                    "crates/litchi-pages/src/package/section_settings.rs",
                    "focused litchi-pages section-settings public API is missing "
                    "Package method edit_section_settings: "
                    "crates/litchi-pages/src/package/section_settings.rs",
                    "focused litchi-pages section-settings public API is missing "
                    "Package method section_settings: "
                    "crates/litchi-pages/src/package/section_settings.rs",
                    "focused litchi-pages section-settings public API is missing "
                    "private package owner source: "
                    "crates/litchi-pages/src/package/section_settings.rs"
                ],
            )

    def test_focused_pages_section_settings_rejects_aliases_and_physical_leaks(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            add_pages_section_settings_canonical_scaffold(root)
            helper = root / boundaries.PAGES_SECTION_SETTINGS_OWNER_HELPER_ROOT / "api.rs"
            helper.parent.mkdir(parents=True)
            helper.write_text(
                "pub type SectionSettingsEdit = Edit;\n"
                "pub fn section_settings(object_id: u64, source_bytes: &[u8], "
                "wire: WireView, archive: Archive, generated: GeneratedProjection, "
                "buffa: BuffaView, prost: prost_types::MessageInfo) {}\n",
                encoding="utf-8",
            )
            lib = root / boundaries.PAGES_SECTION_SETTINGS_EXPORT_SOURCES[0]
            lib.write_text(
                "pub mod section;\n"
                "pub use crate::section::settings::{Settings, Edit};\n",
                encoding="utf-8",
            )
            section = root / boundaries.PAGES_SECTION_SETTINGS_EXPORT_SOURCES[2]
            section.write_text(
                "pub mod settings;\n"
                "pub struct Settings;\n"
                "pub use settings::{Patch, Commit};\n",
                encoding="utf-8",
            )

            violations = boundaries.audit_pages_section_settings_facade_source_topology(
                root
            )

            for fragment in (
                "retains flat alias SectionSettingsEdit",
                "exposes raw identifier object_id",
                "exposes raw source bytes source_bytes",
                "exposes raw byte slice &[u8]",
                "exposes wire type WireView",
                "exposes archive/IWA type Archive",
                "exposes generated type GeneratedProjection",
                "exposes protobuf type BuffaView",
                "exposes protobuf type prost",
                "exposes protobuf type prost_types",
                "exposes public section-settings owner alias",
                "retains root alias Settings",
                "retains root alias Edit",
                "retains root alias Patch",
                "retains root alias Commit",
            ):
                self.assertTrue(
                    any(fragment in item for item in violations),
                    msg=f"missing violation containing {fragment!r}: {violations!r}",
                )

    def test_focused_pages_section_settings_allows_canonical_and_adjacent_api(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            add_pages_section_settings_canonical_scaffold(root)
            semantic = root / boundaries.PAGES_SECTION_SETTINGS_SEMANTIC_SOURCE
            semantic.write_text(
                "pub use crate::package::section_settings::{Commit, Diagnostics, "
                "Edit, Error, LimitKind, Patch, Path, DependencyKind};\n",
                encoding="utf-8",
            )
            owner = root / boundaries.PAGES_SECTION_SETTINGS_OWNER_SOURCE
            owner.write_text(
                "impl Package {\n"
                "pub fn section_settings() {}\n"
                "pub fn edit_section_settings() {}\n"
                "pub fn apply_section_settings() {}\n"
                "}\n"
                "pub(crate) fn resolve_object_id(source_bytes: &[u8], wire: WireView) {}\n",
                encoding="utf-8",
            )
            section = root / boundaries.PAGES_SECTION_SETTINGS_EXPORT_SOURCES[2]
            section.write_text(
                "pub mod pagination;\n"
                "pub mod settings;\n"
                "pub use pagination::{PageNumber, PageNumbering, Pagination, Start};\n"
                "pub struct Settings;\n"
                "pub enum Error { BackgroundPayloadTooLarge }\n",
                encoding="utf-8",
            )

            self.assertEqual(
                boundaries.audit_pages_section_settings_facade_source_topology(root), []
            )


    def test_pages_section_background_boundary_inventories_are_exact(self) -> None:
        self.assertEqual(
            boundaries.RETIRED_IWA_PAGES_SECTION_BACKGROUND_METHODS,
            ("section_background", "set_section_background"),
        )
        self.assertEqual(
            boundaries.RETIRED_IWA_PAGES_SECTION_BACKGROUND_SOURCES,
            (
                Path("crates/litchi-iwa/src/pages/editor/section_background.rs"),
                Path("crates/litchi-iwa/src/pages/editor/section_settings.rs"),
            ),
        )
        self.assertEqual(
            boundaries.PAGES_SECTION_BACKGROUND_CANONICAL_TYPES,
            ("Edit", "Patch", "Commit", "Diagnostics", "Error", "LimitKind", "Path"),
        )
        self.assertEqual(
            boundaries.PAGES_SECTION_BACKGROUND_PACKAGE_METHODS,
            (
                "section_background",
                "edit_section_background",
                "apply_section_background",
            ),
        )
        self.assertEqual(
            boundaries.PAGES_SECTION_BACKGROUND_EDIT_METHODS,
            ("background", "set_solid", "clear", "commit"),
        )
        self.assertTrue(
            {
                "SectionBackground",
                "SectionBackgroundEdit",
                "PagesSectionBackground",
                "PagesSectionBackgroundPatch",
            }
            <= boundaries.PAGES_SECTION_BACKGROUND_FLAT_ALIASES
        )

    def test_retired_iwa_pages_section_background_surface_cannot_return(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            host = root / boundaries.IWA_PAGES_SOURCE_ROOT / "editor/section_background.rs"
            host.parent.mkdir(parents=True)
            host.write_text(
                "pub fn section_background() {}\n"
                "pub fn set_section_background() {}\n"
                "pub(super) fn section_background_payload() {}\n",
                encoding="utf-8",
            )
            retired_settings = root / boundaries.RETIRED_IWA_PAGES_SECTION_BACKGROUND_SOURCES[1]
            retired_settings.write_text("mod preserved {}\n", encoding="utf-8")
            tests = root / boundaries.IWA_PAGES_EDITOR_TEST_SOURCE
            tests.write_text(
                "\n".join(
                    f"fn {name}() {{}}"
                    for name in boundaries.RETIRED_IWA_PAGES_SECTION_BACKGROUND_TESTS
                )
                + "\n",
                encoding="utf-8",
            )
            editor = root / boundaries.IWA_PAGES_EDITOR_SOURCE
            editor.write_text(
                "mod section_background;\npub(crate) mod section_settings;\n",
                encoding="utf-8",
            )
            example = root / boundaries.RETIRED_IWA_PAGES_SECTION_BACKGROUND_EXAMPLE
            example.parent.mkdir(parents=True, exist_ok=True)
            example.write_text("fn main() {}\n", encoding="utf-8")

            violations = boundaries.audit_iwa_pages_section_background_source_topology(
                root
            )

            self.assertEqual(len(violations), 8)
            self.assertEqual(sum("source returned" in item for item in violations), 2)
            self.assertTrue(any("example returned" in item for item in violations))
            for method in boundaries.RETIRED_IWA_PAGES_SECTION_BACKGROUND_METHODS:
                self.assertTrue(
                    any(f"background method {method}:" in item for item in violations)
                )
            for module in boundaries.RETIRED_IWA_PAGES_SECTION_BACKGROUND_MODULES:
                self.assertTrue(
                    any(f"background module {module}:" in item for item in violations)
                )
            for name in boundaries.RETIRED_IWA_PAGES_SECTION_BACKGROUND_TESTS:
                self.assertTrue(any(f"test {name}:" in item for item in violations))
            self.assertFalse(any("section_background_payload" in item for item in violations))

    def test_retired_iwa_pages_section_background_readme_tokens(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            readme = root / boundaries.IWA_PAGES_README
            readme.parent.mkdir(parents=True)
            readme.write_text(
                "pages.section_background(selector);\n"
                "editor\n  .\n  set_section_background(selector, value);\n"
                "other.set_section_background(selector, value);\n"
                "crate::PagesEditor::section_background(selector);\n"
                "set_pages_section_background\n"
                "set_pages_section_background.rs\n",
                encoding="utf-8",
            )

            violations = boundaries.audit_iwa_pages_section_background_source_topology(
                root
            )

            self.assertEqual(sum("README call" in item for item in violations), 4)
            self.assertEqual(
                sum("README example reference" in item for item in violations), 2
            )

    def test_iwa_pages_section_background_policy_retains_adjacent_seams(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            host = root / boundaries.IWA_PAGES_SOURCE_ROOT / "editor/section_wire.rs"
            host.parent.mkdir(parents=True)
            host.write_text(
                "pub(super) fn section_background_payload() {}\n"
                "pub(super) fn set_section_background_payload() {}\n"
                "pub fn section_pagination() {}\n"
                "pub fn edit_section_name() {}\n",
                encoding="utf-8",
            )
            readme = root / boundaries.IWA_PAGES_README
            readme.parent.mkdir(parents=True, exist_ok=True)
            readme.write_text(
                "Use litchi_pages::section::background and edit_section_background.\n",
                encoding="utf-8",
            )

            self.assertEqual(
                boundaries.audit_iwa_pages_section_background_source_topology(root), []
            )

    def test_focused_pages_section_background_requires_each_canonical_type(self) -> None:
        for missing in boundaries.PAGES_SECTION_BACKGROUND_CANONICAL_TYPES:
            with self.subTest(missing=missing):
                with tempfile.TemporaryDirectory() as directory:
                    root = Path(directory)
                    add_pages_section_background_canonical_scaffold(root)
                    semantic = root / boundaries.PAGES_SECTION_BACKGROUND_SEMANTIC_SOURCE
                    semantic.write_text(
                        "".join(
                            f"pub struct {name};\n"
                            for name in boundaries.PAGES_SECTION_BACKGROUND_CANONICAL_TYPES
                            if name != missing
                        )
                        + "impl Edit {\n"
                        + "".join(
                            f"pub fn {method}() {{}}\n"
                            for method in boundaries.PAGES_SECTION_BACKGROUND_EDIT_METHODS
                        )
                        + "}\n",
                        encoding="utf-8",
                    )

                    self.assertEqual(
                        boundaries.audit_pages_section_background_facade_source_topology(
                            root
                        ),
                        [
                            "focused litchi-pages section-background public API is "
                            f"missing canonical section::background type {missing}: "
                            "crates/litchi-pages/src/section/background.rs"
                        ],
                    )

    def test_focused_pages_section_background_requires_nested_private_owner(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            add_pages_section_background_canonical_scaffold(root)
            package = root / boundaries.PAGES_SECTION_BACKGROUND_EXPORT_SOURCES[1]
            package.write_text("pub mod section_background;\n", encoding="utf-8")
            section = root / boundaries.PAGES_SECTION_BACKGROUND_EXPORT_SOURCES[2]
            section.write_text("mod background;\n", encoding="utf-8")

            violations = boundaries.audit_pages_section_background_facade_source_topology(
                root
            )
            self.assertTrue(
                any(
                    "missing canonical section::background module" in item
                    for item in violations
                )
            )
            self.assertTrue(
                any(
                    "exposes duplicate package::section_background module" in item
                    for item in violations
                )
            )

        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            add_pages_section_background_canonical_scaffold(root)
            (root / boundaries.PAGES_SECTION_BACKGROUND_OWNER_SOURCE).unlink()
            self.assertEqual(
                boundaries.audit_pages_section_background_facade_source_topology(root),
                [
                    "focused litchi-pages section-background public API is missing "
                    "Package method apply_section_background: "
                    "crates/litchi-pages/src/package/section_background.rs",
                    "focused litchi-pages section-background public API is missing "
                    "Package method edit_section_background: "
                    "crates/litchi-pages/src/package/section_background.rs",
                    "focused litchi-pages section-background public API is missing "
                    "Package method section_background: "
                    "crates/litchi-pages/src/package/section_background.rs",
                    "focused litchi-pages section-background public API is missing "
                    "private package owner source: "
                    "crates/litchi-pages/src/package/section_background.rs",
                ],
            )

    def test_focused_pages_section_background_rejects_aliases_and_physical_leaks(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            add_pages_section_background_canonical_scaffold(root)
            helper = (
                root / boundaries.PAGES_SECTION_BACKGROUND_OWNER_HELPER_ROOT / "api.rs"
            )
            helper.parent.mkdir(parents=True)
            helper.write_text(
                "pub type SectionBackgroundEdit = Edit;\n"
                "pub fn section_background(object_id: u64, source_bytes: &[u8], "
                "wire: WireView, archive: Archive, generated: GeneratedProjection, "
                "buffa: BuffaView, prost: prost_types::MessageInfo, opaque: Opaque) {}\n",
                encoding="utf-8",
            )
            lib = root / boundaries.PAGES_SECTION_BACKGROUND_EXPORT_SOURCES[0]
            lib.write_text(
                "pub mod section;\n"
                "pub use crate::section::background::{Edit, Patch};\n",
                encoding="utf-8",
            )
            section = root / boundaries.PAGES_SECTION_BACKGROUND_EXPORT_SOURCES[2]
            section.write_text(
                "pub mod background;\npub struct Background;\n"
                "pub use background::{Edit, Patch, Commit};\n",
                encoding="utf-8",
            )

            violations = boundaries.audit_pages_section_background_facade_source_topology(
                root
            )

            for fragment in (
                "retains flat alias SectionBackgroundEdit",
                "exposes raw identifier object_id",
                "exposes raw source bytes source_bytes",
                "exposes raw byte slice &[u8]",
                "exposes wire type WireView",
                "exposes archive/IWA type Archive",
                "exposes generated type GeneratedProjection",
                "exposes protobuf type BuffaView",
                "exposes protobuf type prost",
                "exposes protobuf type prost_types",
                "exposes archive/IWA type Opaque",
                "exposes public section-background owner alias",
                "retains root alias Edit",
                "retains root alias Patch",
                "retains root alias Commit",
            ):
                self.assertTrue(
                    any(fragment in item for item in violations),
                    msg=f"missing violation containing {fragment!r}: {violations!r}",
                )

    def test_focused_pages_section_background_allows_canonical_and_adjacent_api(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            add_pages_section_background_canonical_scaffold(root)
            semantic = root / boundaries.PAGES_SECTION_BACKGROUND_SEMANTIC_SOURCE
            semantic.write_text(
                "pub use crate::package::section_background::{Commit, Diagnostics, "
                "Edit, Error, LimitKind, Patch, Path};\n"
                "impl Edit {\n"
                "pub fn background() {}\n"
                "pub fn set_solid() {}\n"
                "pub fn clear() {}\n"
                "pub fn commit() {}\n"
                "}\n",
                encoding="utf-8",
            )
            owner = root / boundaries.PAGES_SECTION_BACKGROUND_OWNER_SOURCE
            owner.write_text(
                "impl Package {\n"
                "pub fn section_background() {}\n"
                "pub fn edit_section_background() {}\n"
                "pub fn apply_section_background() {}\n"
                "}\n"
                "pub(crate) fn resolve_object_id(source_bytes: &[u8], wire: WireView) {}\n",
                encoding="utf-8",
            )
            section = root / boundaries.PAGES_SECTION_BACKGROUND_EXPORT_SOURCES[2]
            section.write_text(
                "pub mod background;\npub struct Background;\n"
                "pub mod settings;\npub struct Settings;\n",
                encoding="utf-8",
            )

            self.assertEqual(
                boundaries.audit_pages_section_background_facade_source_topology(root),
                [],
            )


    def test_numbers_table_cells_read_boundary_inventories_are_exact(self) -> None:
        self.assertEqual(
            boundaries.NUMBERS_TABLE_CELLS_SEMANTIC_SOURCE,
            Path("crates/litchi-numbers/src/table/cells.rs"),
        )
        self.assertEqual(
            boundaries.NUMBERS_TABLE_CELLS_OWNER_SOURCE,
            Path("crates/litchi-numbers/src/package/table_cells.rs"),
        )
        self.assertEqual(
            boundaries.NUMBERS_TABLE_CELLS_CANONICAL_TYPES,
            ("State", "Storage", "Error", "LimitKind", "Path"),
        )
        self.assertEqual(
            boundaries.NUMBERS_TABLE_CELLS_PACKAGE_METHODS,
            ("table_cell", "table_cells"),
        )
        for deferred in (
            "Input",
            "Change",
            "Edit",
            "Patch",
            "Commit",
            "Diagnostics",
            "DependencyKind",
        ):
            self.assertNotIn(deferred, boundaries.NUMBERS_TABLE_CELLS_CANONICAL_TYPES)
        for deferred in ("edit_table_cells", "apply_table_cells"):
            self.assertNotIn(deferred, boundaries.NUMBERS_TABLE_CELLS_PACKAGE_METHODS)

    def test_focused_numbers_table_cells_read_requires_each_type(self) -> None:
        for missing in boundaries.NUMBERS_TABLE_CELLS_CANONICAL_TYPES:
            with self.subTest(missing=missing):
                with tempfile.TemporaryDirectory() as directory:
                    root = Path(directory)
                    add_numbers_table_cells_read_scaffold(root)
                    semantic = root / boundaries.NUMBERS_TABLE_CELLS_SEMANTIC_SOURCE
                    retained = [
                        name
                        for name in boundaries.NUMBERS_TABLE_CELLS_CANONICAL_TYPES
                        if name != missing
                    ]
                    semantic.write_text(
                        "".join(f"pub struct {name};\n" for name in retained),
                        encoding="utf-8",
                    )

                    self.assertEqual(
                        boundaries.audit_numbers_table_cells_facade_source_topology(
                            root
                        ),
                        [
                            "focused litchi-numbers table-cells read API is missing "
                            f"canonical table::cells type {missing}: "
                            "crates/litchi-numbers/src/table/cells.rs"
                        ],
                    )

    def test_focused_numbers_table_cells_read_requires_each_package_method(
        self,
    ) -> None:
        for missing in boundaries.NUMBERS_TABLE_CELLS_PACKAGE_METHODS:
            with self.subTest(missing=missing):
                with tempfile.TemporaryDirectory() as directory:
                    root = Path(directory)
                    add_numbers_table_cells_read_scaffold(root)
                    owner = root / boundaries.NUMBERS_TABLE_CELLS_OWNER_SOURCE
                    owner.write_text(
                        "impl Package {\n"
                        + "".join(
                            f"pub fn {name}() {{}}\n"
                            for name in boundaries.NUMBERS_TABLE_CELLS_PACKAGE_METHODS
                            if name != missing
                        )
                        + "}\n",
                        encoding="utf-8",
                    )

                    self.assertEqual(
                        boundaries.audit_numbers_table_cells_facade_source_topology(
                            root
                        ),
                        [
                            "focused litchi-numbers table-cells read API is missing "
                            f"canonical Package::{missing} method: "
                            "crates/litchi-numbers/src/package/table_cells.rs"
                        ],
                    )

    def test_focused_numbers_table_cells_read_requires_private_owner(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            add_numbers_table_cells_read_scaffold(root)
            package = root / boundaries.NUMBERS_TABLE_CELLS_EXPORT_SOURCES[1]
            package.write_text("pub mod table_cells;\n", encoding="utf-8")
            table = root / boundaries.NUMBERS_TABLE_CELLS_EXPORT_SOURCES[2]
            table.write_text("mod cells;\n", encoding="utf-8")

            violations = boundaries.audit_numbers_table_cells_facade_source_topology(
                root
            )

            self.assertTrue(
                any("missing canonical table::cells module" in item for item in violations)
            )
            self.assertTrue(
                any(
                    "exposes duplicate package::table_cells module" in item
                    for item in violations
                )
            )

        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            add_numbers_table_cells_read_scaffold(root)
            (root / boundaries.NUMBERS_TABLE_CELLS_OWNER_SOURCE).unlink()
            violations = boundaries.audit_numbers_table_cells_facade_source_topology(
                root
            )
            self.assertTrue(
                any("missing private package owner source" in item for item in violations)
            )
            self.assertEqual(
                sum("canonical Package::" in item for item in violations),
                2,
            )

    def test_focused_numbers_table_cells_read_rejects_aliases_and_physical_leaks(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            add_numbers_table_cells_read_scaffold(root)
            semantic = root / boundaries.NUMBERS_TABLE_CELLS_SEMANTIC_SOURCE
            semantic.write_text(
                semantic.read_text(encoding="utf-8")
                + "pub type TableCellState = State;\n",
                encoding="utf-8",
            )
            owner = root / boundaries.NUMBERS_TABLE_CELLS_OWNER_SOURCE
            owner.write_text(
                "impl Package {\n"
                "pub fn table_cell(object_id: u64, source_bytes: &[u8], "
                "bnc: BncCell, codec: NumbersTableCellStorageCodec, "
                "buffa: BuffaCellView, generated: GeneratedProjection, "
                "prost: prost_types::MessageInfo) {}\n"
                "pub fn table_cells() {}\n"
                "}\n",
                encoding="utf-8",
            )
            lib = root / boundaries.NUMBERS_TABLE_CELLS_EXPORT_SOURCES[0]
            lib.write_text(
                "pub mod table;\n"
                "pub use crate::table::cells::{State, Storage};\n",
                encoding="utf-8",
            )

            violations = boundaries.audit_numbers_table_cells_facade_source_topology(
                root
            )

            for fragment in (
                "retains flat alias TableCellState",
                "exposes raw identifier object_id",
                "exposes raw source bytes source_bytes",
                "exposes raw byte slice &[u8]",
                "exposes wire/BNC type BncCell",
                "exposes protobuf type NumbersTableCellStorageCodec",
                "exposes protobuf type BuffaCellView",
                "exposes generated type GeneratedProjection",
                "exposes protobuf type prost",
                "exposes protobuf type prost_types",
                "exposes public table-cells owner alias",
                "retains root alias State",
                "retains root alias Storage",
            ):
                self.assertTrue(
                    any(fragment in item for item in violations),
                    msg=f"missing violation containing {fragment!r}: {violations!r}",
                )

    def test_focused_numbers_table_cells_read_allows_semantic_variants_and_types(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            add_numbers_table_cells_read_scaffold(root)
            semantic = root / boundaries.NUMBERS_TABLE_CELLS_SEMANTIC_SOURCE
            semantic.write_text(
                "pub use crate::cell::{FiniteF64, Value};\n"
                "pub use crate::cell::data_format::DataFormat;\n"
                "pub use crate::table::{CellPosition, CellRange, Dimensions};\n"
                "pub struct State;\n"
                "pub enum Storage { Empty, Stored }\n"
                "pub struct Error;\n"
                "pub struct LimitKind;\n"
                "pub struct Path;\n"
                "pub enum SemanticDependency { CellStorage, FormulaCache }\n",
                encoding="utf-8",
            )

            self.assertEqual(
                boundaries.audit_numbers_table_cells_facade_source_topology(root),
                [],
            )

    def test_numbers_table_cells_mutation_boundary_inventories_are_exact(self) -> None:
        self.assertEqual(
            boundaries.NUMBERS_TABLE_CELLS_MUTATION_TYPES,
            (
                "Input",
                "Change",
                "Edit",
                "Patch",
                "Commit",
                "Diagnostics",
                "DependencyKind",
            ),
        )
        self.assertEqual(
            boundaries.NUMBERS_TABLE_CELLS_FULL_CANONICAL_TYPES,
            (
                "Input",
                "Change",
                "State",
                "Storage",
                "Edit",
                "Patch",
                "Commit",
                "Diagnostics",
                "Error",
                "LimitKind",
                "Path",
                "DependencyKind",
            ),
        )
        self.assertNotIn("Plan", boundaries.NUMBERS_TABLE_CELLS_FULL_CANONICAL_TYPES)
        self.assertEqual(
            boundaries.NUMBERS_TABLE_CELLS_MUTATION_PACKAGE_METHODS,
            ("edit_table_cells", "apply_table_cells"),
        )
        self.assertEqual(
            boundaries.NUMBERS_TABLE_CELLS_FULL_PACKAGE_METHODS,
            ("table_cell", "table_cells", "edit_table_cells", "apply_table_cells"),
        )
        self.assertEqual(
            boundaries.NUMBERS_TABLE_CELLS_MUTATION_OWNER_SOURCE,
            Path("crates/litchi-numbers/src/package/table_cell_edit.rs"),
        )
        self.assertEqual(
            boundaries.RETIRED_IWA_NUMBERS_TABLE_CELL_SOURCE_INVENTORY,
            (
                Path("crates/litchi-iwa/src/numbers/editor/semantic/table.rs"),
                Path("crates/litchi-iwa/src/numbers/editor/model.rs"),
                Path("crates/litchi-iwa/src/numbers/editor/table_cells.rs"),
                Path("crates/litchi-iwa/src/numbers/editor/tests.rs"),
                Path("crates/litchi-iwa/examples/edit_numbers_cell.rs"),
            ),
        )
        self.assertEqual(
            boundaries.RETIRED_IWA_NUMBERS_TABLE_CELL_METHODS,
            ("set_cell", "set_cells", "clear_cell"),
        )
        self.assertEqual(
            boundaries.RETIRED_IWA_NUMBERS_TABLE_CELL_MODEL_HELPERS,
            ("set_cell_in_package", "set_cells_in_package"),
        )
        self.assertEqual(
            boundaries.RETIRED_IWA_NUMBERS_TABLE_CELL_BATCH_HELPERS,
            ("apply_numbers",),
        )
        self.assertEqual(
            boundaries.RETIRED_IWA_NUMBERS_TABLE_CELL_TESTS,
            (
                "semantic_edits_round_trip_through_public_reader",
                "cell_batch_roundtrips_mixed_values_and_clear",
                "cell_batch_refreshes_formula_chain_from_final_state",
                "cell_batch_rejects_invalid_inputs_transactionally",
                "failed_edit_is_transactional",
                "cell_edits_keep_sparse_row_headers_in_lockstep",
                "source_created_large_table_allocates_sparse_tiles_for_batch_writes",
                "rich_text_cell_updates_preserve_the_payload_reference",
                "shared_rich_text_cell_update_uses_copy_on_write",
                "replacing_rich_text_releases_list_and_payload_objects",
                "segmented_string_entries_round_trip_and_remain_interned",
                "segmented_shared_rich_text_uses_copy_on_write_and_cleans_up",
                "formula_cells_can_be_cleared_with_refcount_cleanup",
                "cell_write_refreshes_transitive_formula_caches_in_dependency_order",
                "cell_write_rejects_unsupported_impacted_formula_transactionally",
            ),
        )
        self.assertEqual(
            boundaries.IWA_NUMBERS_TABLE_CELL_TEST_FIXTURE_SOURCE,
            Path("crates/litchi-iwa/src/numbers/editor.rs"),
        )
        self.assertEqual(
            boundaries.IWA_NUMBERS_TABLE_CELL_TEST_FIXTURE_HELPERS,
            ("test_set_cell", "test_set_cells", "test_clear_cell"),
        )

    def test_focused_numbers_table_cells_mutation_requires_each_type(self) -> None:
        for missing in boundaries.NUMBERS_TABLE_CELLS_MUTATION_TYPES:
            with self.subTest(missing=missing):
                with tempfile.TemporaryDirectory() as directory:
                    root = Path(directory)
                    add_numbers_table_cells_mutation_scaffold(root)
                    semantic = root / boundaries.NUMBERS_TABLE_CELLS_SEMANTIC_SOURCE
                    semantic.write_text(
                        "".join(
                            f"pub struct {name};\n"
                            for name in boundaries.NUMBERS_TABLE_CELLS_FULL_CANONICAL_TYPES
                            if name != missing
                        ),
                        encoding="utf-8",
                    )

                    self.assertEqual(
                        boundaries.audit_numbers_table_cells_mutation_facade_source_topology(
                            root
                        ),
                        [
                            "focused litchi-numbers table-cells mutation API is missing "
                            f"canonical table::cells type {missing}: "
                            "crates/litchi-numbers/src/table/cells.rs"
                        ],
                    )

    def test_focused_numbers_table_cells_mutation_requires_each_package_method(
        self,
    ) -> None:
        for missing in boundaries.NUMBERS_TABLE_CELLS_MUTATION_PACKAGE_METHODS:
            with self.subTest(missing=missing):
                with tempfile.TemporaryDirectory() as directory:
                    root = Path(directory)
                    add_numbers_table_cells_mutation_scaffold(root)
                    owner = root / boundaries.NUMBERS_TABLE_CELLS_MUTATION_OWNER_SOURCE
                    owner.write_text(
                        "impl Package {\n"
                        + "".join(
                            f"pub fn {name}() {{}}\n"
                            for name in (
                                boundaries.NUMBERS_TABLE_CELLS_MUTATION_PACKAGE_METHODS
                            )
                            if name != missing
                        )
                        + "}\n",
                        encoding="utf-8",
                    )

                    self.assertEqual(
                        boundaries.audit_numbers_table_cells_mutation_facade_source_topology(
                            root
                        ),
                        [
                            "focused litchi-numbers table-cells mutation API is missing "
                            f"canonical Package::{missing} method: "
                            "crates/litchi-numbers/src/package/table_cell_edit.rs"
                        ],
                    )

    def test_focused_numbers_table_cells_mutation_requires_private_owner(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            add_numbers_table_cells_mutation_scaffold(root)
            package = root / boundaries.NUMBERS_TABLE_CELLS_EXPORT_SOURCES[1]
            package.write_text(
                "pub(crate) mod table_cells;\npub mod table_cell_edit;\n",
                encoding="utf-8",
            )

            violations = (
                boundaries.audit_numbers_table_cells_mutation_facade_source_topology(
                    root
                )
            )

            self.assertEqual(len(violations), 1)
            self.assertIn("exposes package::table_cell_edit module", violations[0])

        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            add_numbers_table_cells_mutation_scaffold(root)
            (root / boundaries.NUMBERS_TABLE_CELLS_MUTATION_OWNER_SOURCE).unlink()
            violations = (
                boundaries.audit_numbers_table_cells_mutation_facade_source_topology(
                    root
                )
            )
            self.assertTrue(
                any("missing private package owner source" in item for item in violations)
            )
            self.assertEqual(
                sum("canonical Package::" in item for item in violations),
                2,
            )

    def test_focused_numbers_table_cells_mutation_rejects_public_leaks(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            add_numbers_table_cells_mutation_scaffold(root)
            semantic = root / boundaries.NUMBERS_TABLE_CELLS_SEMANTIC_SOURCE
            semantic.write_text(
                semantic.read_text(encoding="utf-8")
                + "pub type TableCellEdit = Edit;\n",
                encoding="utf-8",
            )
            owner = root / boundaries.NUMBERS_TABLE_CELLS_MUTATION_OWNER_SOURCE
            owner.write_text(
                "impl Package {\n"
                "pub fn edit_table_cells(object_id: u64, source_bytes: &[u8], "
                "wire: WireView, bnc: BncCell, codec: NumbersTableCellStorageCodec, "
                "buffa: BuffaCellView, generated: GeneratedProjection, "
                "prost: prost_types::MessageInfo) {}\n"
                "pub fn apply_table_cells() {}\n"
                "}\n",
                encoding="utf-8",
            )
            lib = root / boundaries.NUMBERS_TABLE_CELLS_EXPORT_SOURCES[0]
            lib.write_text(
                "pub mod table;\n"
                "pub use crate::table::cells::{Edit, Patch, Commit};\n",
                encoding="utf-8",
            )

            violations = (
                boundaries.audit_numbers_table_cells_mutation_facade_source_topology(
                    root
                )
            )

            for fragment in (
                "retains flat alias TableCellEdit",
                "exposes raw identifier object_id",
                "exposes raw source bytes source_bytes",
                "exposes raw byte slice &[u8]",
                "exposes wire/BNC type WireView",
                "exposes wire/BNC type BncCell",
                "exposes protobuf type NumbersTableCellStorageCodec",
                "exposes protobuf type BuffaCellView",
                "exposes generated type GeneratedProjection",
                "exposes protobuf type prost",
                "exposes protobuf type prost_types",
                "exposes public mutation-owner alias",
                "retains root alias Edit",
                "retains root alias Patch",
                "retains root alias Commit",
            ):
                self.assertTrue(
                    any(fragment in item for item in violations),
                    msg=f"missing violation containing {fragment!r}: {violations!r}",
                )

    def test_focused_numbers_table_cells_mutation_allows_private_plan_and_variants(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            add_numbers_table_cells_mutation_scaffold(root)
            semantic = root / boundaries.NUMBERS_TABLE_CELLS_SEMANTIC_SOURCE
            semantic.write_text(
                "pub struct Input;\npub struct Change;\n"
                "pub struct State;\npub struct Storage;\n"
                "pub struct Edit;\npub struct Patch;\npub struct Commit;\n"
                "pub struct Diagnostics;\npub struct Error;\n"
                "pub struct LimitKind;\npub struct Path;\n"
                "pub enum DependencyKind { CellStorage, FormulaCache }\n"
                "pub(crate) struct Plan;\n",
                encoding="utf-8",
            )

            self.assertEqual(
                boundaries.audit_numbers_table_cells_mutation_facade_source_topology(
                    root
                ),
                [],
            )

    def test_iwa_numbers_table_cell_mutation_rejects_exact_retired_surface(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            sources = {
                boundaries.RETIRED_IWA_NUMBERS_TABLE_CELL_EDITOR_SOURCE: (
                    boundaries.RETIRED_IWA_NUMBERS_TABLE_CELL_METHODS
                ),
                boundaries.RETIRED_IWA_NUMBERS_TABLE_CELL_MODEL_SOURCE: (
                    boundaries.RETIRED_IWA_NUMBERS_TABLE_CELL_MODEL_HELPERS
                ),
                boundaries.RETIRED_IWA_NUMBERS_TABLE_CELL_BATCH_SOURCE: (
                    boundaries.RETIRED_IWA_NUMBERS_TABLE_CELL_BATCH_HELPERS
                ),
                boundaries.RETIRED_IWA_NUMBERS_TABLE_CELL_TEST_SOURCE: (
                    boundaries.RETIRED_IWA_NUMBERS_TABLE_CELL_TESTS
                ),
            }
            for path, names in sources.items():
                source = root / path
                source.parent.mkdir(parents=True, exist_ok=True)
                source.write_text(
                    "".join(f"fn {name}() {{}}\n" for name in names),
                    encoding="utf-8",
                )
            example = root / boundaries.RETIRED_IWA_NUMBERS_TABLE_CELL_EXAMPLE
            example.parent.mkdir(parents=True, exist_ok=True)
            example.write_text("fn main() {}\n", encoding="utf-8")

            violations = (
                boundaries.audit_iwa_numbers_table_cell_mutation_source_topology(root)
            )

            self.assertEqual(len(violations), 22)
            for name in (
                *boundaries.RETIRED_IWA_NUMBERS_TABLE_CELL_METHODS,
                *boundaries.RETIRED_IWA_NUMBERS_TABLE_CELL_MODEL_HELPERS,
                *boundaries.RETIRED_IWA_NUMBERS_TABLE_CELL_BATCH_HELPERS,
                *boundaries.RETIRED_IWA_NUMBERS_TABLE_CELL_TESTS,
            ):
                self.assertTrue(
                    any(f" {name}:" in violation for violation in violations),
                    msg=f"missing retirement violation for {name!r}: {violations!r}",
                )
            self.assertTrue(any("example returned" in item for item in violations))

    def test_iwa_numbers_table_cell_mutation_allows_retained_host_helpers(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            retained_sources = {
                boundaries.RETIRED_IWA_NUMBERS_TABLE_CELL_EDITOR_SOURCE: (
                    "set_cell_comment",
                    "clear_cell_comment",
                    "set_cell_conditional_highlighting",
                ),
                boundaries.RETIRED_IWA_NUMBERS_TABLE_CELL_MODEL_SOURCE: (
                    "set_attached_cell_in_package",
                    "set_attached_cells_in_package",
                ),
                boundaries.RETIRED_IWA_NUMBERS_TABLE_CELL_BATCH_SOURCE: (
                    "is_empty",
                    "len",
                    "collect",
                    "apply_attached",
                    "into_formula_cache_coordinates",
                ),
                boundaries.RETIRED_IWA_NUMBERS_TABLE_CELL_TEST_SOURCE: (
                    "retained_table_cell_formula_cache_test",
                ),
                Path("crates/litchi-iwa/src/pages/editor/tables/semantic.rs"): (
                    "set_cell",
                    "set_cells",
                ),
                Path("crates/litchi-iwa/src/keynote/editor/slide_tables.rs"): (
                    "clear_cell",
                    "apply_numbers",
                ),
                Path("crates/litchi-iwa/src/numbers/editor/package.rs"): (
                    "set_table_cell_in_package",
                ),
            }
            for path, names in retained_sources.items():
                source = root / path
                source.parent.mkdir(parents=True, exist_ok=True)
                source.write_text(
                    "".join(f"fn {name}() {{}}\n" for name in names),
                    encoding="utf-8",
                )
            fixture = root / boundaries.IWA_NUMBERS_TABLE_CELL_TEST_FIXTURE_SOURCE
            fixture.parent.mkdir(parents=True, exist_ok=True)
            fixture.write_text(
                "".join(
                    f"#[cfg(test)]\npub(crate) fn {name}() {{}}\n"
                    for name in boundaries.IWA_NUMBERS_TABLE_CELL_TEST_FIXTURE_HELPERS
                ),
                encoding="utf-8",
            )

            self.assertEqual(
                boundaries.audit_iwa_numbers_table_cell_mutation_source_topology(root),
                [],
            )

    def test_iwa_numbers_table_cell_fixture_helpers_stay_test_only(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            fixture = root / boundaries.IWA_NUMBERS_TABLE_CELL_TEST_FIXTURE_SOURCE
            fixture.parent.mkdir(parents=True, exist_ok=True)
            fixture.write_text(
                "pub(crate) fn test_set_cell() {}\n"
                "#[cfg(test)]\npub fn test_set_cells() {}\n"
                "#[cfg(test)]\npub(crate) fn test_clear_cell() {}\n",
                encoding="utf-8",
            )
            example = root / boundaries.IWA_NUMBERS_EXAMPLE_ROOT / "fixture_leak.rs"
            example.parent.mkdir(parents=True, exist_ok=True)
            example.write_text(
                "fn main() { crate::numbers::editor::test_clear_cell(); }\n",
                encoding="utf-8",
            )

            violations = (
                boundaries.audit_iwa_numbers_table_cell_mutation_source_topology(root)
            )

            self.assertEqual(len(violations), 3)
            self.assertTrue(
                any("private #[cfg(test)] test_set_cell:" in item for item in violations)
            )
            self.assertTrue(
                any("private #[cfg(test)] test_set_cells:" in item for item in violations)
            )
            self.assertTrue(
                any("example calls test-only" in item for item in violations)
            )


if __name__ == "__main__":
    unittest.main()
