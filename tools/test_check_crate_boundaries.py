from __future__ import annotations

import copy
import json
import unittest
from dataclasses import replace

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
    return boundaries.Snapshot(
        packages=policy.packages,
        manifests=frozenset(),
        edges={edge: ("kind=normal, optional=false, target=*, rename=-",) for edge in edges},
        dependencies=frozen_dependencies,
        normal_dependencies=frozen_dependencies,
        features=features,
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
        features["litchi"] |= frozenset({"ole"})
        snapshot = replace(snapshot, features=features)

        violations = boundaries.audit_snapshot(snapshot, self.policy)

        self.assertIn("retired litchi facade features returned: ole", violations)

    def test_violations_have_deterministic_order(self) -> None:
        snapshot = valid_snapshot(self.policy)
        snapshot = replace(
            snapshot,
            packages=snapshot.packages | frozenset({"z-new", "a-new"}),
        )

        violations = boundaries.audit_snapshot(snapshot, self.policy)

        self.assertEqual(violations, sorted(violations))
        self.assertIn("workspace packages lack topology policy: a-new, z-new", violations)


if __name__ == "__main__":
    unittest.main()
