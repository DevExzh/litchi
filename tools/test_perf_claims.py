"""Standard-library tests for the v1 performance claim registry."""

from __future__ import annotations

import copy
import hashlib
import json
import tempfile
import unittest
from pathlib import Path

from tools import check_perf_claims


REPO_ROOT = Path(__file__).resolve().parents[1]
REGISTRY_PATH = REPO_ROOT / "docs" / "performance" / "claim-registry-v1.json"


def load_seed() -> dict:
    return check_perf_claims.load_json(REGISTRY_PATH)


class ClaimRegistryStructuralTests(unittest.TestCase):
    def test_seed_registry_is_structurally_valid(self) -> None:
        registry = load_seed()
        _, evidence, claims = check_perf_claims.validate_registry(
            registry, repo_root=REPO_ROOT
        )
        self.assertEqual(set(evidence), {"abba-0248", "abba-0249", "abba-0250", "abba-0251"})
        self.assertEqual(
            set(claims),
            {
                "claim-0248-cfb-streaming",
                "claim-0249-ods-known-change",
                "claim-0250-zip-ordering",
                "claim-0251-xlsx-xml-borrowed",
            },
        )

    def test_seed_registry_cli_structural_mode(self) -> None:
        status, messages = check_perf_claims.lint_registry(
            REGISTRY_PATH,
            repo_root=REPO_ROOT,
            evidence_root=None,
            mode="structural",
        )
        self.assertEqual(status, 0, messages)

    def test_canonical_registry_digest_is_deterministic(self) -> None:
        registry = load_seed()
        first = hashlib.sha256(check_perf_claims.canonical_bytes(registry)).hexdigest()
        second = hashlib.sha256(check_perf_claims.canonical_bytes(copy.deepcopy(registry))).hexdigest()
        self.assertEqual(first, second)
        self.assertEqual(len(first), 64)

    def test_duplicate_json_keys_are_rejected(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "duplicate.json"
            path.write_text('{"schema_version": 1, "schema_version": 1}\n', encoding="utf-8")
            with self.assertRaises(check_perf_claims.ClaimInputError):
                check_perf_claims.load_json(path)

    def test_unknown_registry_key_is_rejected(self) -> None:
        registry = load_seed()
        registry["unexpected"] = True
        with self.assertRaises(check_perf_claims.ClaimInputError):
            check_perf_claims.validate_registry(registry, repo_root=REPO_ROOT)

    def test_uppercase_digest_is_rejected(self) -> None:
        registry = load_seed()
        registry["evidence"][0]["manifest"]["sha256"] = registry["evidence"][0]["manifest"]["sha256"].upper()
        with self.assertRaises(check_perf_claims.ClaimInputError):
            check_perf_claims.validate_registry(registry, repo_root=REPO_ROOT)

    def test_traversal_document_reference_is_rejected(self) -> None:
        registry = load_seed()
        registry["claims"][0]["documentation"] = ["../AGENTS.md"]
        with self.assertRaises(check_perf_claims.ClaimInputError):
            check_perf_claims.validate_registry(registry, repo_root=REPO_ROOT)

    def test_landed_claim_requires_verified_evidence(self) -> None:
        registry = load_seed()
        claim = registry["claims"][0]
        claim["status"] = "landed"
        claim["code_state"] = "landed"
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "registry.json"
            path.write_text(json.dumps(registry), encoding="utf-8")
            status, messages = check_perf_claims.lint_registry(
                path,
                repo_root=REPO_ROOT,
                evidence_root=None,
                mode="structural",
            )
        self.assertEqual(status, 2, messages)


class ResourceReportShapeTests(unittest.TestCase):
    METRICS = (
        "heaptrack.allocation_calls",
        "heaptrack.allocated_bytes",
        "heaptrack.peak_heap_bytes",
        "heaptrack.peak_rss_bytes",
        "time.max_rss_kib",
    )

    def make_report(self) -> dict:
        legs = []
        for label, variant, binary, revision in (
            ("A1", "control", "a" * 64, "1" * 40),
            ("B1", "candidate", "b" * 64, "2" * 40),
            ("B2", "candidate", "b" * 64, "2" * 40),
            ("A2", "control", "a" * 64, "1" * 40),
        ):
            legs.append(
                {
                    "leg": label,
                    "variant": variant,
                    "binary_identity": {"binary_sha256": binary},
                    "harness_report": {
                        "environment": {
                            "git_worktree_dirty": False,
                            "git_revision": revision,
                        }
                    },
                }
            )
        metrics = {}
        for metric in self.METRICS:
            metrics[metric] = {
                "paired": {
                    "A1_control_to_B1_candidate": {
                        "control": 100.0,
                        "candidate": 102.0,
                        "relative_delta_percent": 2.0,
                        "status": "observed",
                    },
                    "A2_control_to_B2_candidate": {
                        "control": 100.0,
                        "candidate": 101.0,
                        "relative_delta_percent": 1.0,
                        "status": "observed",
                    },
                }
            }
        return {
            "schema_version": 1,
            "abba_schema_version": 1,
            "validation": {
                "status": "validated",
                "control_revision": "1" * 40,
                "candidate_revision": "2" * 40,
            },
            "legs": legs,
            "statistics": {"status": "observed", "metrics": metrics},
        }

    def test_valid_resource_shape(self) -> None:
        report = self.make_report()
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            path = root / "resource.json"
            path.write_text(json.dumps(report), encoding="utf-8")
            digest = hashlib.sha256(path.read_bytes()).hexdigest()
            evidence = {
                "id": "resource-test",
                "kind": "resource_abba_report",
                "relative_path": "resource.json",
                "sha256": digest,
                "schema_version": 1,
                "abba_schema_version": 1,
            }
            claim = {"resource": {"metrics": list(self.METRICS)}}
            check_perf_claims.verify_resource_report(
                evidence,
                claim,
                evidence_root=root,
                resource_policy={"max_regression_percent": 5},
                latency_result=None,
            )

    def test_withheld_resource_metric_is_not_claimable(self) -> None:
        report = self.make_report()
        report["statistics"]["metrics"][self.METRICS[0]]["paired"]["A1_control_to_B1_candidate"]["status"] = "not_measured"
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            path = root / "resource.json"
            path.write_text(json.dumps(report), encoding="utf-8")
            evidence = {
                "id": "resource-test",
                "kind": "resource_abba_report",
                "relative_path": "resource.json",
                "sha256": hashlib.sha256(path.read_bytes()).hexdigest(),
                "schema_version": 1,
                "abba_schema_version": 1,
            }
            with self.assertRaises(check_perf_claims.ClaimPolicyError):
                check_perf_claims.verify_resource_report(
                    evidence,
                    {"resource": {"metrics": list(self.METRICS)}},
                    evidence_root=root,
                    resource_policy={"max_regression_percent": 5},
                    latency_result=None,
                )


if __name__ == "__main__":
    unittest.main()
