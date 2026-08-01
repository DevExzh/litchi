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
    features = {name: frozenset() for name in policy.packages}
    features["litchi-core"] = frozenset(item.name for item in policy.core_feature_debt)
    return boundaries.Snapshot(
        packages=policy.packages,
        manifests=frozenset(),
        edges={edge: ("kind=normal, optional=false, target=*, rename=-",) for edge in edges},
        dependencies={name: frozenset(items) for name, items in dependencies.items()},
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
        edge = boundaries.Edge("litchi-xlsx", "litchi-ole")
        edges = dict(snapshot.edges)
        edges[edge] = ("kind=dev, optional=true, target=cfg(unix), rename=legacy",)
        dependencies = dict(snapshot.dependencies)
        dependencies[edge.dependent] |= frozenset({edge.dependency})
        snapshot = replace(snapshot, edges=edges, dependencies=dependencies)

        violations = boundaries.audit_snapshot(snapshot, self.policy)

        self.assertIn(
            "unclassified internal edge litchi-xlsx -> litchi-ole "
            "(kind=dev, optional=true, target=cfg(unix), rename=legacy)",
            violations,
        )

    def test_resolved_migration_edge_requires_policy_cleanup(self) -> None:
        snapshot = valid_snapshot(self.policy)
        edge = self.policy.migration_debt[0].edge
        edges = dict(snapshot.edges)
        del edges[edge]
        dependencies = dict(snapshot.dependencies)
        dependencies[edge.dependent] -= frozenset({edge.dependency})
        snapshot = replace(snapshot, edges=edges, dependencies=dependencies)

        violations = boundaries.audit_snapshot(snapshot, self.policy)

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

    def test_migration_host_edge_cannot_be_marked_canonical(self) -> None:
        raw = copy.deepcopy(self.raw_policy)
        raw["packages"]["litchi-ole"] = ["litchi-core"]
        raw["migration_debt"] = [
            item
            for item in raw["migration_debt"]
            if not (
                item["dependent"] == "litchi-ole"
                and item["dependency"] == "litchi-core"
            )
        ]

        with self.assertRaisesRegex(
            boundaries.PolicyError,
            "migration-host edges must be debt, not canonical: litchi-ole -> litchi-core",
        ):
            boundaries.parse_policy(raw)

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
