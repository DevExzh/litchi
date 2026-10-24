"""Tests for the fixed W8 original-to-final projection policy."""

from __future__ import annotations

import copy
import hashlib
import importlib.util
import json
from pathlib import Path
import unittest

from tools import w8_projection_policy as policy


class ProjectionPolicyTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.before = policy._load_fixed_projection("before")
        cls.after = policy._load_fixed_projection("after")

    def test_reviewed_policy_and_fixed_pair(self) -> None:
        value = policy.load_reviewed_policy()
        self.assertEqual(value, policy._policy_summary())
        proof = policy.validate_projection_pair(self.before, self.after)
        self.assertEqual(proof["registry_package_count"], 518)
        self.assertTrue(proof["registry_identities_checksums_dependencies_equal"])
        self.assertEqual(len(proof["local_edge_differences"]), 6)

    def test_policy_bytes_are_registered_and_canonical(self) -> None:
        path = policy.ROOT / policy.POLICY_RELATIVE
        payload = path.read_bytes()
        self.assertEqual(hashlib.sha256(payload).hexdigest(), policy.REVIEWED_POLICY_SHA256)
        self.assertEqual(payload, policy.canonical(json.loads(payload)))

    def test_registry_checksum_tamper_is_rejected(self) -> None:
        changed = copy.deepcopy(self.after)
        package = next(item for item in changed["lock"]["packages"] if item.get("source", "").startswith("registry+"))
        package["checksum"] = "0" * 64
        with self.assertRaises(policy.ProjectionPolicyError):
            policy.validate_projection_pair(self.before, changed)

    def test_registry_dependency_tamper_is_rejected(self) -> None:
        changed = copy.deepcopy(self.after)
        package = next(item for item in changed["lock"]["packages"] if item.get("source", "").startswith("registry+"))
        package["dependencies"] = list(package["dependencies"]) + ["invented-edge"]
        with self.assertRaises(policy.ProjectionPolicyError):
            policy.validate_projection_pair(self.before, changed)

    def test_local_delta_tamper_is_rejected(self) -> None:
        changed = copy.deepcopy(self.after)
        package = next(item for item in changed["lock"]["packages"] if item["name"] == "litchi-xlsx")
        package["dependencies"] = list(package["dependencies"]) + ["invented-edge"]
        with self.assertRaises(policy.ProjectionPolicyError):
            policy.validate_projection_pair(self.before, changed)

    def test_feature_and_workload_tamper_are_rejected(self) -> None:
        for field, value in (("features", {"package": "litchi-ops", "selected": ["default"], "all_features": True}), ("workload", {"route_count": 108})):
            changed = copy.deepcopy(self.after)
            changed[field] = value
            with self.subTest(field=field), self.assertRaises(policy.ProjectionPolicyError):
                policy.validate_projection_pair(self.before, changed)

    def test_generic_policy_path_cannot_be_replaced(self) -> None:
        with self.assertRaises(policy.ProjectionPolicyError):
            policy.load_reviewed_policy(Path("/tmp/another-policy.json"))

    def test_generic_runner_identity_contract_remains_strict(self) -> None:
        path = Path(__file__).resolve().parents[1] / "docs/unified-ops/results/0362h/run.py"
        spec = importlib.util.spec_from_file_location("w8_run_contract", path)
        self.assertIsNotNone(spec)
        self.assertIsNotNone(spec.loader)
        runner = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(runner)
        before = {
            "dependency_projection_sha256": "a",
            "lock": {"sha256": "b"},
            "probe": {"sha256": "c"},
            "fixture": {"sha256": "d"},
        }
        self.assertEqual(runner.shared_identity(before, copy.deepcopy(before)), {
            "dependency_projection_sha256": "a",
            "lock_sha256": "b",
            "probe_sha256": "c",
            "fixture_sha256": "d",
        })
        changed = copy.deepcopy(before)
        changed["lock"]["sha256"] = "changed"
        with self.assertRaises(RuntimeError):
            runner.shared_identity(before, changed)


if __name__ == "__main__":
    unittest.main()
