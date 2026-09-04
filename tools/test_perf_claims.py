"""Standard-library tests for the v1 performance claim registry."""

from __future__ import annotations

import copy
import hashlib
import io
import json
import shutil
import subprocess
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from tools import check_perf_claims, perf_abba_summary


REPO_ROOT = Path(__file__).resolve().parents[1]
REGISTRY_PATH = REPO_ROOT / "docs" / "performance" / "claim-registry-v1.json"
SCHEMA_PATH = REPO_ROOT / "docs" / "performance" / "claim-registry-v1.schema.json"


def load_seed() -> dict:
    return check_perf_claims.load_json(REGISTRY_PATH)


class ClaimRegistryStructuralTests(unittest.TestCase):
    def test_seed_registry_is_structurally_valid(self) -> None:
        registry = load_seed()
        _, evidence, claims = check_perf_claims.validate_registry(
            registry, repo_root=REPO_ROOT
        )
        self.assertEqual(
            set(evidence),
            {
                "abba-0248",
                "abba-0249",
                "abba-0250",
                "abba-0251",
                "resource-0251",
                "abba-0268-xls-owned-source",
                "abba-0269-xlsx-repeated-store-cache",
            },
        )
        self.assertEqual(
            set(claims),
            {
                "claim-0248-cfb-streaming",
                "claim-0249-ods-known-change",
                "claim-0250-zip-ordering",
                "claim-0251-xlsx-xml-borrowed",
                "claim-0268-xls-owned-source",
                "claim-0269-xlsx-repeated-store-cache",
            },
        )
        claim = claims["claim-0251-xlsx-xml-borrowed"]["value"]
        self.assertEqual(claim["status"], "landed")
        self.assertEqual(claim["code_state"], "landed")
        self.assertEqual(claim["resource_guardrail"]["evidence_id"], "resource-0251")

    def test_seed_registry_cli_strict_mode(self) -> None:
        status, messages = check_perf_claims.lint_registry(
            REGISTRY_PATH,
            repo_root=REPO_ROOT,
            evidence_root=REPO_ROOT,
            mode="strict",
        )
        self.assertEqual(status, 0, messages)

    def test_strict_mode_accepts_landed_latency_only_claim(self) -> None:
        for evidence_id, claim_id in (
            ("abba-0268-xls-owned-source", "claim-0268-xls-owned-source"),
            (
                "abba-0269-xlsx-repeated-store-cache",
                "claim-0269-xlsx-repeated-store-cache",
            ),
        ):
            with self.subTest(claim_id=claim_id):
                registry = load_seed()
                registry["evidence"] = [
                    next(
                        item for item in registry["evidence"] if item["id"] == evidence_id
                    )
                ]
                registry["claims"] = [
                    next(
                        item for item in registry["claims"] if item["id"] == claim_id
                    )
                ]
                self.assertNotIn("resource_guardrail", registry["claims"][0])
                with tempfile.TemporaryDirectory() as directory:
                    path = Path(directory) / "registry.json"
                    path.write_text(json.dumps(registry), encoding="utf-8")
                    with patch.object(
                        check_perf_claims,
                        "verify_abba_package",
                        return_value={},
                    ) as verify_abba:
                        status, messages = check_perf_claims.lint_registry(
                            path,
                            repo_root=REPO_ROOT,
                            evidence_root=REPO_ROOT,
                            mode="strict",
                        )
                self.assertEqual(status, 0, messages)
                verify_abba.assert_called_once()

    def test_landed_declared_resource_guardrail_still_requires_evidence(self) -> None:
        registry = load_seed()
        claim = next(
            item
            for item in registry["claims"]
            if item["id"] == "claim-0268-xls-owned-source"
        )
        claim["resource_guardrail"] = {
            "required": True,
            "metrics": ["time.max_rss_kib"],
        }
        with self.assertRaisesRegex(
            check_perf_claims.ClaimPolicyError,
            "landed claim requires a resource evidence_id",
        ):
            check_perf_claims.validate_registry(registry, repo_root=REPO_ROOT)

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


class ClaimRegistryMetricExtensionTests(unittest.TestCase):
    def partial_landed_registry(self) -> dict:
        registry = load_seed()
        claim = next(
            item
            for item in registry["claims"]
            if item["id"] == "claim-0268-xls-owned-source"
        )
        latency = claim["latency_evidence"]
        moved = latency["accepted_statistics"][-1]
        latency["accepted_statistics"] = latency["accepted_statistics"][:-1]
        latency["accepted_cells"] -= 1
        latency["adverse_both_cells"] = 1
        latency["adverse_both_statistics"] = [moved]
        latency["metric_profile"] = "elapsed_ns"
        return registry

    def test_schema_is_additive_and_closes_metric_profile_vocabulary(self) -> None:
        schema = json.loads(SCHEMA_PATH.read_text(encoding="utf-8"))
        self.assertIn("Structural registry shape only", schema["description"])
        self.assertIn("Cross-field claim counts", schema["description"])
        self.assertIn("tools/check_perf_claims.py", schema["description"])
        latency = schema["$defs"]["latencyEvidence"]
        self.assertEqual(
            set(latency["properties"]),
            {
                "evidence_id",
                "metric_profile",
                "allowed_statistics",
                "accepted_cells",
                "adverse_both_cells",
                "accepted_statistics",
                "adverse_both_statistics",
            },
        )
        self.assertEqual(
            schema["$defs"]["metricProfile"]["enum"],
            ["elapsed_ns", "publication_ns"],
        )
        self.assertEqual(
            set(schema["$defs"]["latencyEvidence"]["required"]),
            {"evidence_id", "allowed_statistics", "accepted_cells", "adverse_both_cells"},
        )

    def test_landed_claim_can_declare_partial_accepted_and_adverse_cells(self) -> None:
        registry = self.partial_landed_registry()
        _, _, claims = check_perf_claims.validate_registry(registry, repo_root=REPO_ROOT)
        parsed = claims["claim-0268-xls-owned-source"]
        self.assertEqual(parsed["metric_profile"], "elapsed_ns")
        self.assertEqual(parsed["latency"]["accepted_cells"], 7)
        self.assertEqual(parsed["adverse_cells"][0]["statistic"], "p99")

        implicit_registry = self.partial_landed_registry()
        implicit_claim = next(
            item
            for item in implicit_registry["claims"]
            if item["id"] == "claim-0268-xls-owned-source"
        )
        implicit_claim["latency_evidence"].pop("metric_profile")
        _, _, implicit_claims = check_perf_claims.validate_registry(
            implicit_registry, repo_root=REPO_ROOT
        )
        self.assertEqual(
            implicit_claims["claim-0268-xls-owned-source"]["metric_profile"],
            "elapsed_ns",
        )

    def test_future_change_entries_are_not_added_to_seed_registry(self) -> None:
        registry = load_seed()
        self.assertFalse(
            any(
                item["change_id"].startswith(("0401-", "0402-"))
                for item in registry["claims"]
            )
        )

    def test_0402_docs_match_supported_profile_and_intentional_no_entry(self) -> None:
        paths = (
            REPO_ROOT / "docs" / "performance" / "ADR_COMPLIANCE.md",
            REPO_ROOT / "docs" / "performance" / "CRUD_COVERAGE.md",
            REPO_ROOT
            / "docs"
            / "performance"
            / "changes"
            / "0402-opc-overlay-decoder-reuse.md",
        )
        for path in paths:
            with self.subTest(path=path):
                normalized = " ".join(path.read_text(encoding="utf-8").split()).lower()
                self.assertNotIn("cannot represent", normalized)
                self.assertNotIn("schema cannot express", normalized)
                self.assertIn("publication_ns", normalized)
                self.assertIn("now supports", normalized)
                self.assertIn("partial accepted/adverse-cell adjudication", normalized)
                self.assertTrue(
                    "has no 0402 claim entry" in normalized
                    or "adds no 0402 registry claim entry" in normalized
                )

    def test_registry_boundary_docs_name_real_allocator_and_inventory_checks(self) -> None:
        claim_doc = (REPO_ROOT / "docs" / "performance" / "claim-registry-v1.md").read_text(
            encoding="utf-8"
        )
        checker_source = (REPO_ROOT / "tools" / "check_perf_claims.py").read_text(
            encoding="utf-8"
        )
        for text in (claim_doc, checker_source):
            normalized = " ".join(text.split()).lower()
            self.assertIn("tools/validate_opc_overlay_allocator_abba.py", normalized)
            self.assertIn("evidence-manifest.json", normalized)
            self.assertNotIn("subject to their dedicated validator", normalized)
            self.assertNotIn(
                "must be validated separately by their existing validator", normalized
            )
        self.assertTrue(
            (REPO_ROOT / "tools" / "validate_opc_overlay_allocator_abba.py").is_file()
        )

    def test_unsupported_metric_profile_is_rejected(self) -> None:
        registry = load_seed()
        claim = next(item for item in registry["claims"] if item["id"] == "claim-0268-xls-owned-source")
        claim["latency_evidence"]["metric_profile"] = "arbitrary_metric"
        with self.assertRaisesRegex(check_perf_claims.ClaimInputError, "unsupported"):
            check_perf_claims.validate_registry(registry, repo_root=REPO_ROOT)

    def test_status_and_code_state_mismatch_is_rejected(self) -> None:
        registry = load_seed()
        claim = next(item for item in registry["claims"] if item["id"] == "claim-0268-xls-owned-source")
        claim["status"] = "landed"
        claim["code_state"] = "not_landed"
        with self.assertRaisesRegex(check_perf_claims.ClaimPolicyError, "status/code_state disagree"):
            check_perf_claims.validate_registry(registry, repo_root=REPO_ROOT)

    def test_adverse_cell_count_and_set_overlap_fail_closed(self) -> None:
        registry = self.partial_landed_registry()
        claim = next(item for item in registry["claims"] if item["id"] == "claim-0268-xls-owned-source")
        claim["latency_evidence"]["adverse_both_cells"] = 2
        with self.assertRaisesRegex(check_perf_claims.ClaimPolicyError, "adverse_both_statistics"):
            check_perf_claims.validate_registry(registry, repo_root=REPO_ROOT)

        registry = self.partial_landed_registry()
        claim = next(item for item in registry["claims"] if item["id"] == "claim-0268-xls-owned-source")
        claim["latency_evidence"]["adverse_both_statistics"] = [
            copy.deepcopy(claim["latency_evidence"]["accepted_statistics"][0])
        ]
        with self.assertRaisesRegex(check_perf_claims.ClaimPolicyError, "overlap"):
            check_perf_claims.validate_registry(registry, repo_root=REPO_ROOT)

    def test_landed_partial_claim_requires_explicit_adverse_cells(self) -> None:
        registry = load_seed()
        claim = next(item for item in registry["claims"] if item["id"] == "claim-0268-xls-owned-source")
        claim["latency_evidence"]["metric_profile"] = "elapsed_ns"
        claim["latency_evidence"]["adverse_both_cells"] = 1
        with self.assertRaisesRegex(check_perf_claims.ClaimPolicyError, "requires adverse_both_statistics"):
            check_perf_claims.validate_registry(registry, repo_root=REPO_ROOT)

    def test_publication_profile_extracts_nested_metric_cells(self) -> None:
        summary = check_perf_claims.load_json(
            REPO_ROOT / "docs" / "performance" / "results" / "change-0402" / "summary.json"
        )
        profile = check_perf_claims._resolve_metric_profile(
            {"metric_profile": "publication_ns"}, "test.latency_evidence"
        )
        accepted, adverse = check_perf_claims._canonical_summary_cells(
            summary,
            location="test.summary",
            metric_profile=profile,
        )
        self.assertEqual(len(accepted), 13)
        self.assertEqual(adverse, set())

    def test_publication_profile_strict_package_recomputes_without_registry_entry(self) -> None:
        package = REPO_ROOT / "docs" / "performance" / "results" / "change-0402"
        manifest_path = package / "0402-opc-source-overlay-abba-manifest.json"
        summary_path = package / "summary.json"
        manifest = check_perf_claims.load_json(manifest_path)
        summary = check_perf_claims.load_json(summary_path)
        accepted: list[dict[str, str]] = []
        corpora: list[dict[str, str]] = []
        seen_corpora: set[str] = set()
        for row in summary["results"]:
            corpus = row["corpus"]
            if corpus["name"] not in seen_corpora:
                corpora.append(
                    {"name": corpus["name"], "archive_sha256": corpus["archive_sha256"]}
                )
                seen_corpora.add(corpus["name"])
            accepted.extend(
                {
                    "case": row["case"],
                    "corpus": corpus["name"],
                    "statistic": statistic,
                }
                for statistic in row["publication_ns"]["accepted_statistics"]
            )
        evidence = {
            "id": "abba-publication-test",
            "relative_path": "docs/performance/results/change-0402",
            "manifest": {
                "path": manifest_path.name,
                "sha256": check_perf_claims.sha256_file(manifest_path)[0],
                "schema_version": 1,
            },
            "summary": {
                "path": "summary.json",
                "sha256": check_perf_claims.sha256_file(summary_path)[0],
                "canonical_sha256": check_perf_claims.canonical_sha256(summary),
                "schema_version": 1,
            },
        }
        claim = {
            "id": "claim-publication-test",
            "change_id": manifest["change_id"],
            "status": "landed",
            "code_state": "landed",
            "scope": {
                "format": "OPC/ZIP",
                "selectors": [summary["results"][0]["case"]],
                "corpora": corpora,
            },
            "latency": {
                "metric_profile": "publication_ns",
                "allowed_statistics": list(check_perf_claims.STATISTICS),
                "accepted_cells": len(accepted),
                "adverse_both_cells": 0,
            },
            "accepted_cells": accepted,
            "adverse_cells": [],
        }
        policy = load_seed()["policies"]["latency-abba-v1"]
        result = check_perf_claims.verify_abba_package(
            evidence,
            claim,
            evidence_root=REPO_ROOT,
            policy=policy,
        )
        self.assertEqual(result["result_count"], 9)

        injected_profile_claim = copy.deepcopy(claim)
        injected_profile_claim["latency"]["metric_profile"] = "elapsed_ns"
        injected_profile_claim["_metric_spec"] = check_perf_claims._resolve_metric_profile(
            {"metric_profile": "publication_ns"}, "injected"
        )
        with self.assertRaisesRegex(check_perf_claims.ClaimInputError, "unsupported schema"):
            check_perf_claims.verify_abba_package(
                evidence,
                injected_profile_claim,
                evidence_root=REPO_ROOT,
                policy=policy,
            )

    def test_elapsed_profile_strict_package_accepts_adverse_cell(self) -> None:
        package = REPO_ROOT / "docs" / "performance" / "results" / "change-0401"
        manifest_path = package / "0401-xlsx-selected-numeric-elision-abba-manifest.json"
        summary_path = package / "summary.json"
        manifest = check_perf_claims.load_json(manifest_path)
        summary = check_perf_claims.load_json(summary_path)
        row = summary["results"][0]
        corpus = row["corpus"]
        accepted = [
            {"case": row["case"], "corpus": corpus["name"], "statistic": statistic}
            for statistic in row["elapsed_ns"]["accepted_statistics"]
        ]
        adverse = [
            {"case": row["case"], "corpus": corpus["name"], "statistic": statistic}
            for statistic in row["elapsed_ns"]["adverse_both_statistics"]
        ]
        evidence = {
            "id": "abba-elapsed-test",
            "relative_path": "docs/performance/results/change-0401",
            "manifest": {
                "path": manifest_path.name,
                "sha256": check_perf_claims.sha256_file(manifest_path)[0],
                "schema_version": 1,
            },
            "summary": {
                "path": "summary.json",
                "sha256": check_perf_claims.sha256_file(summary_path)[0],
                "canonical_sha256": check_perf_claims.canonical_sha256(summary),
                "schema_version": 1,
            },
        }
        claim = {
            "id": "claim-elapsed-test",
            "change_id": manifest["change_id"],
            "status": "landed",
            "code_state": "landed",
            "scope": {
                "format": "XLSX",
                "selectors": [row["case"]],
                "corpora": [
                    {"name": corpus["name"], "archive_sha256": corpus["archive_sha256"]}
                ],
            },
            "latency": {
                "metric_profile": "elapsed_ns",
                "allowed_statistics": list(check_perf_claims.STATISTICS),
                "accepted_cells": len(accepted),
                "adverse_both_cells": len(adverse),
            },
            "accepted_cells": accepted,
            "adverse_cells": adverse,
        }
        policy = load_seed()["policies"]["latency-abba-v1"]
        result = check_perf_claims.verify_abba_package(
            evidence,
            claim,
            evidence_root=REPO_ROOT,
            policy=policy,
        )
        self.assertEqual(result["result_count"], 1)


class StrictLatencyRecomputationTests(unittest.TestCase):
    EVIDENCE_ID = "abba-0251"

    def load_fixture(self) -> tuple[dict, dict[str, dict], dict]:
        registry = load_seed()
        evidence = next(item for item in registry["evidence"] if item["id"] == self.EVIDENCE_ID)
        summary_path = (
            REPO_ROOT
            / "docs"
            / "performance"
            / "results"
            / "0251-xlsx-xml-borrowed-20260821"
            / "summary.json"
        )
        summary = check_perf_claims.load_json(summary_path)
        package = summary_path.parent
        reports: dict[str, dict] = {}
        for role in check_perf_claims.ABBA_ROLES:
            raw = subprocess.check_output(
                ["zstd", "-q", "-d", "-c", str(package / f"{role}.json.zst")]
            )
            reports[role] = json.loads(raw)
        return registry, reports, summary

    def recompute(self, reports: dict[str, dict], summary: dict) -> None:
        registry = load_seed()
        try:
            projections = {
                role: perf_abba_summary._project_report(
                    reports[role],
                    role,
                    profile=perf_abba_summary.detect_report_profile(reports[role], role),
                )
                for role in check_perf_claims.ABBA_ROLES
            }
            canonical = perf_abba_summary._summarize_projected_reports(
                projections,
                drift_ceilings=registry["policies"]["latency-abba-v1"]["drift_ceiling_percent"],
            )
        except Exception as error:
            raise check_perf_claims.ClaimInputError(str(error)) from error
        if check_perf_claims.canonical_sha256(canonical) != check_perf_claims.canonical_sha256(summary):
            raise check_perf_claims.ClaimInputError(
                "summary differs from canonical raw-report recomputation"
            )

    def test_raw_sample_tamper_is_rejected_despite_recomputed_flag(self) -> None:
        _, reports, summary = self.load_fixture()
        reports["a1"]["results"][0]["elapsed_ns"]["samples"][-1] += 1
        self.assertTrue(summary["verification"]["statistics_recomputed_from_samples"])
        with self.assertRaises(check_perf_claims.ClaimInputError):
            self.recompute(reports, summary)

    def test_summary_statistic_tamper_is_rejected(self) -> None:
        _, reports, summary = self.load_fixture()
        summary = copy.deepcopy(summary)
        summary["results"][0]["elapsed_ns"]["legs_ns"]["a1"]["p50"] += 1
        self.assertTrue(summary["verification"]["statistics_recomputed_from_samples"])
        with self.assertRaises(check_perf_claims.ClaimInputError):
            self.recompute(reports, summary)

    def test_summary_accepted_set_tamper_is_rejected(self) -> None:
        _, reports, summary = self.load_fixture()
        summary = copy.deepcopy(summary)
        elapsed = summary["results"][0]["elapsed_ns"]
        elapsed["accepted_statistics"] = ["p50"]
        self.assertTrue(summary["verification"]["statistics_recomputed_from_samples"])
        with self.assertRaises(check_perf_claims.ClaimInputError):
            self.recompute(reports, summary)


class ResourceReportShapeTests(unittest.TestCase):
    METRICS = (
        "heaptrack.allocation_calls",
        "heaptrack.allocated_bytes",
        "heaptrack.temporary_allocations",
        "heaptrack.peak_heap_bytes",
        "heaptrack.peak_rss_bytes",
        "time.max_rss_kib",
    )

    def make_report(self) -> dict:
        def artifact(seed: str) -> dict:
            digest = hashlib.sha256(seed.encode()).hexdigest()
            return {"bytes": 1, "present": True, "retained": True, "sha256": digest}

        def run(seed: str) -> dict:
            return {
                "returncode": 0,
                "timed_out": False,
                "stdout": artifact(f"{seed}-stdout"),
                "stderr": artifact(f"{seed}-stderr"),
            }

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
                    "binary_identity": {"binary_sha256": binary, "label": variant},
                    "harness_identity": {
                        "leg": label,
                        "variant": variant,
                        "git_revision": revision,
                        "git_worktree_dirty": False,
                        "environment": {"git_revision": revision},
                        "tool": {
                            "name": "litchi-perf-baseline",
                            "version": "0.1.0",
                            "profile": "release",
                        },
                    },
                    # Keep the old envelope present as a compatibility-shape
                    # fixture; strict verification binds the richer identity
                    # above when both are available.
                    "harness_report": {
                        "environment": {
                            "git_worktree_dirty": False,
                            "git_revision": revision,
                        }
                    },
                }
            )
        metrics = {}
        values_by_leg = {
            "A1": 100.0,
            "B1": 102.0,
            "B2": 101.0,
            "A2": 100.0,
        }
        for metric in self.METRICS:
            metrics[metric] = {
                "control": {
                    "status": "observed",
                    "count": 2,
                    "mean": 100.0,
                    "median": 100.0,
                    "minimum": 100.0,
                    "maximum": 100.0,
                    "overflow_fields": [],
                },
                "candidate": {
                    "status": "observed",
                    "count": 2,
                    "mean": 101.5,
                    "median": 101.5,
                    "minimum": 101.0,
                    "maximum": 102.0,
                    "overflow_fields": [],
                },
                "paired": {
                    "A1_control_to_B1_candidate": {
                        "control": 100.0,
                        "candidate": 102.0,
                        "execution_order": "A1, B1",
                        "ratio_candidate_to_control": 1.02,
                        "relative_delta_percent": 2.0,
                        "status": "observed",
                    },
                    "A2_control_to_B2_candidate": {
                        "control": 100.0,
                        "candidate": 101.0,
                        "execution_order": "B2, A2",
                        "ratio_candidate_to_control": 1.01,
                        "relative_delta_percent": 1.0,
                        "status": "observed",
                    },
                },
                "values_by_leg": dict(values_by_leg),
            }
        resource_labels = ("A1", "B1", "B2", "A2")
        for index, leg in enumerate(legs):
            heaptrack_values = {
                metric.split(".", 1)[1]: values_by_leg[resource_labels[index]]
                for metric in self.METRICS
                if metric.startswith("heaptrack.")
            }
            time_values = {
                metric.split(".", 1)[1]: values_by_leg[resource_labels[index]]
                for metric in self.METRICS
                if metric.startswith("time.")
            }
            time_values.update(
                {
                    field: values_by_leg[resource_labels[index]]
                    for field in check_perf_claims._TIME_RESOURCE_FIELDS
                }
            )
            heaptrack_values.update(
                {
                    field: values_by_leg[resource_labels[index]]
                    for field in check_perf_claims._HEAPTRACK_RESOURCE_FIELDS
                }
            )
            capture = artifact(f"{resource_labels[index]}-capture")
            print_artifact = artifact(f"{resource_labels[index]}-print")
            print_run = run(f"{resource_labels[index]}-heaptrack-print")
            print_run["stdout"] = copy.deepcopy(print_artifact)
            leg["heaptrack"] = {
                "status": "ok",
                "harness": artifact(f"{resource_labels[index]}-harness"),
                "harness_identity": {
                    "status": "validated",
                    "sha256": ("a" if resource_labels[index] in {"A1", "A2"} else "b") * 64,
                },
                "run": run(f"{resource_labels[index]}-heaptrack"),
                "capture": capture,
                "captures": [copy.deepcopy(capture)],
                "print": {
                    "status": "ok",
                    "artifact": print_artifact,
                    "run": print_run,
                    "parsed": {
                        "status": "ok",
                        "artifact": copy.deepcopy(print_artifact),
                        "histogram_artifact": artifact(
                            f"{resource_labels[index]}-histogram"
                        ),
                        **heaptrack_values,
                    },
                },
            }
            leg["time"] = {
                "status": "ok",
                "run": run(f"{resource_labels[index]}-time"),
                "parsed": {
                    "status": "ok",
                    "artifact": artifact(f"{resource_labels[index]}-time-parsed"),
                    "expected_fields": list(check_perf_claims._TIME_RESOURCE_FIELDS),
                    **time_values,
                },
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

    def assert_rejected(self, report: dict, metrics: list[str] | None = None) -> None:
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
            with self.assertRaises(check_perf_claims.ClaimInputError):
                check_perf_claims.verify_resource_report(
                    evidence,
                    {"resource": {"metrics": metrics or list(self.METRICS)}},
                    evidence_root=root,
                    resource_policy={"max_regression_percent": 5},
                    latency_result=None,
                )

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

    def test_values_by_leg_mismatch_is_rejected(self) -> None:
        report = self.make_report()
        metric = self.METRICS[0]
        report["statistics"]["metrics"][metric]["values_by_leg"]["A1"] = 101.0
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
            with self.assertRaises(check_perf_claims.ClaimInputError):
                check_perf_claims.verify_resource_report(
                    evidence,
                    {"resource": {"metrics": [metric]}},
                    evidence_root=root,
                    resource_policy={"max_regression_percent": 5},
                    latency_result=None,
                )

    def test_paired_delta_mismatch_is_rejected(self) -> None:
        report = self.make_report()
        metric = self.METRICS[0]
        report["statistics"]["metrics"][metric]["paired"]["A1_control_to_B1_candidate"][
            "relative_delta_percent"
        ] = 99.0
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
            with self.assertRaises(check_perf_claims.ClaimInputError):
                check_perf_claims.verify_resource_report(
                    evidence,
                    {"resource": {"metrics": [metric]}},
                    evidence_root=root,
                    resource_policy={"max_regression_percent": 5},
                    latency_result=None,
                )

    def test_missing_resource_source_is_rejected(self) -> None:
        report = self.make_report()
        metric = self.METRICS[0]
        del report["statistics"]["metrics"][metric]["values_by_leg"]
        for leg in report["legs"]:
            leg.pop("heaptrack", None)
            leg.pop("time", None)
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
            with self.assertRaises(check_perf_claims.ClaimInputError):
                check_perf_claims.verify_resource_report(
                    evidence,
                    {"resource": {"metrics": [metric]}},
                    evidence_root=root,
                    resource_policy={"max_regression_percent": 5},
                    latency_result=None,
                )

    def test_declared_only_resource_source_is_rejected(self) -> None:
        report = self.make_report()
        metric = self.METRICS[0]
        for leg in report["legs"]:
            leg.pop("heaptrack", None)
            leg.pop("time", None)
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
            with self.assertRaises(check_perf_claims.ClaimInputError):
                check_perf_claims.verify_resource_report(
                    evidence,
                    {"resource": {"metrics": [metric]}},
                    evidence_root=root,
                    resource_policy={"max_regression_percent": 5},
                    latency_result=None,
                )

    def test_unsupported_resource_source_is_rejected(self) -> None:
        report = self.make_report()
        metric = "unsupported.metric"
        report["statistics"]["metrics"][metric] = copy.deepcopy(
            report["statistics"]["metrics"][self.METRICS[0]]
        )
        del report["statistics"]["metrics"][metric]["values_by_leg"]
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
            with self.assertRaises(check_perf_claims.ClaimInputError):
                check_perf_claims.verify_resource_report(
                    evidence,
                    {"resource": {"metrics": [metric]}},
                    evidence_root=root,
                    resource_policy={"max_regression_percent": 5},
                    latency_result=None,
                )

    def test_failed_time_run_cannot_retain_stale_parser_value(self) -> None:
        report = self.make_report()
        report["legs"][0]["time"]["run"]["returncode"] = 1
        self.assert_rejected(report, ["time.max_rss_kib"])

    def test_timed_out_time_run_cannot_retain_stale_parser_value(self) -> None:
        report = self.make_report()
        report["legs"][0]["time"]["run"]["timed_out"] = True
        self.assert_rejected(report, ["time.max_rss_kib"])

    def test_failed_heaptrack_capture_run_cannot_retain_stale_parser_value(self) -> None:
        report = self.make_report()
        report["legs"][1]["heaptrack"]["run"]["returncode"] = 1
        self.assert_rejected(report, ["heaptrack.allocated_bytes"])

    def test_failed_heaptrack_print_run_cannot_retain_stale_parser_value(self) -> None:
        report = self.make_report()
        report["legs"][1]["heaptrack"]["print"]["run"]["returncode"] = 1
        self.assert_rejected(report, ["heaptrack.allocated_bytes"])

    def test_failed_outer_resource_status_cannot_retain_stale_parser_value(self) -> None:
        report = self.make_report()
        report["legs"][2]["heaptrack"]["status"] = "failed"
        self.assert_rejected(report, ["heaptrack.allocated_bytes"])

    def test_failed_print_status_cannot_retain_stale_parser_value(self) -> None:
        report = self.make_report()
        report["legs"][2]["heaptrack"]["print"]["status"] = "failed"
        self.assert_rejected(report, ["heaptrack.allocated_bytes"])

    def test_unretained_time_parser_artifact_is_rejected(self) -> None:
        report = self.make_report()
        report["legs"][0]["time"]["parsed"]["artifact"]["retained"] = False
        self.assert_rejected(report, ["time.max_rss_kib"])

    def test_time_parser_schema_fields_are_bound(self) -> None:
        report = self.make_report()
        report["legs"][0]["time"]["parsed"]["expected_fields"] = ["max_rss_kib"]
        self.assert_rejected(report, ["time.max_rss_kib"])

    def test_mismatched_heaptrack_parser_artifact_is_rejected(self) -> None:
        report = self.make_report()
        report["legs"][0]["heaptrack"]["print"]["parsed"]["artifact"]["sha256"] = "c" * 64
        self.assert_rejected(report, ["heaptrack.allocated_bytes"])

    def test_heaptrack_print_output_identity_is_bound(self) -> None:
        report = self.make_report()
        report["legs"][0]["heaptrack"]["print"]["run"]["stdout"]["sha256"] = "c" * 64
        self.assert_rejected(report, ["heaptrack.allocated_bytes"])

    def test_mismatched_capture_artifact_is_rejected(self) -> None:
        report = self.make_report()
        report["legs"][0]["heaptrack"]["captures"][0]["sha256"] = "c" * 64
        self.assert_rejected(report, ["heaptrack.allocated_bytes"])

    def test_leg_variant_and_revision_are_bound_to_validation(self) -> None:
        rebound_variant = self.make_report()
        rebound_variant["legs"][1]["variant"] = "control"
        self.assert_rejected(rebound_variant, ["time.max_rss_kib"])

        rebound_revision = self.make_report()
        rebound_revision["legs"][1]["harness_identity"]["git_revision"] = "1" * 40
        self.assert_rejected(rebound_revision, ["time.max_rss_kib"])

        rebound_environment = self.make_report()
        rebound_environment["legs"][1]["harness_identity"]["environment"]["git_revision"] = "1" * 40
        self.assert_rejected(rebound_environment, ["time.max_rss_kib"])

    def test_binary_and_harness_profiles_are_stable_and_distinct(self) -> None:
        binary_drift = self.make_report()
        binary_drift["legs"][3]["binary_identity"]["mode_bits"] = 1
        self.assert_rejected(binary_drift, ["time.max_rss_kib"])

        tool_drift = self.make_report()
        tool_drift["legs"][1]["harness_identity"]["tool"]["profile"] = "debug"
        self.assert_rejected(tool_drift, ["time.max_rss_kib"])

        harness_drift = self.make_report()
        harness_drift["legs"][3]["heaptrack"]["harness_identity"]["sha256"] = "b" * 64
        self.assert_rejected(harness_drift, ["time.max_rss_kib"])

        candidate_rebound = self.make_report()
        candidate_rebound["legs"][1]["binary_identity"]["binary_sha256"] = "a" * 64
        self.assert_rejected(candidate_rebound, ["time.max_rss_kib"])


class _FakeZstdProcess:
    def __init__(self, stdout: bytes, *, returncode: int = 0, stderr: bytes = b"") -> None:
        self.stdout = io.BytesIO(stdout)
        self.stderr = io.BytesIO(stderr)
        self.returncode = returncode
        self.wait_calls = 0
        self.kill_calls = 0

    def __enter__(self) -> _FakeZstdProcess:
        return self

    def __exit__(self, exc_type: object, exc_value: object, traceback: object) -> bool:
        self.wait()
        self.stdout.close()
        self.stderr.close()
        return False

    def wait(self) -> int:
        self.wait_calls += 1
        return self.returncode

    def kill(self) -> None:
        self.kill_calls += 1


class DecompressJsonProcessCleanupTests(unittest.TestCase):
    def test_popen_pipes_close_on_success_and_parse_failure(self) -> None:
        cases = ((b'{"ok":true}', False), (b"not-json", True))
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "report.json.zst"
            for payload, should_fail in cases:
                with self.subTest(should_fail=should_fail):
                    process = _FakeZstdProcess(payload)
                    with patch.object(check_perf_claims.subprocess, "Popen", return_value=process):
                        if should_fail:
                            with self.assertRaises(check_perf_claims.ClaimInputError):
                                check_perf_claims._decompress_json(source, location="test/report")
                        else:
                            report, digest, size = check_perf_claims._decompress_json(source, location="test/report")
                            self.assertEqual(report, {"ok": True})
                            self.assertEqual(digest, hashlib.sha256(payload).hexdigest())
                            self.assertEqual(size, len(payload))
                    self.assertGreaterEqual(process.wait_calls, 1)
                    self.assertTrue(process.stdout.closed)
                    self.assertTrue(process.stderr.closed)

    def test_member_decompressed_ceiling_is_enforced_while_streaming(self) -> None:
        process = _FakeZstdProcess(b"{}" + b"x")
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "report.json.zst"
            with patch.object(check_perf_claims.subprocess, "Popen", return_value=process):
                with self.assertRaisesRegex(check_perf_claims.ClaimInputError, "ceiling"):
                    check_perf_claims._decompress_json(
                        source,
                        location="test/report",
                        max_bytes=2,
                    )
        self.assertEqual(process.kill_calls, 1)


class StrictPackageProjectionHardeningTests(unittest.TestCase):
    EVIDENCE_ID = "abba-0251"
    CLAIM_ID = "claim-0251-xlsx-xml-borrowed"

    def setUp(self) -> None:
        registry = load_seed()
        self.policy = registry["policies"]["latency-abba-v1"]
        self.evidence = copy.deepcopy(
            next(item for item in registry["evidence"] if item["id"] == self.EVIDENCE_ID)
        )
        _, evidence_by_id, claims_by_id = check_perf_claims.validate_registry(
            registry, repo_root=REPO_ROOT
        )
        self.evidence = copy.deepcopy(evidence_by_id[self.EVIDENCE_ID])
        parsed = claims_by_id[self.CLAIM_ID]
        self.claim = {
            **parsed["value"],
            "scope": parsed["scope"],
            "latency": parsed["latency"],
            "accepted_cells": parsed["accepted_cells"],
            "resource": parsed["resource"],
        }
        self.tempdir = tempfile.TemporaryDirectory()
        self.root = Path(self.tempdir.name)
        source = REPO_ROOT / self.evidence["relative_path"]
        shutil.copytree(source, self.root / source.name)
        self.evidence["relative_path"] = source.name

    def tearDown(self) -> None:
        self.tempdir.cleanup()

    def verify(self) -> None:
        check_perf_claims.verify_abba_package(
            self.evidence,
            self.claim,
            evidence_root=self.root,
            policy=self.policy,
        )

    def test_derived_summary_tamper_is_rejected_after_rebinding_hashes(self) -> None:
        package = self.root / self.evidence["relative_path"]
        summary_path = package / self.evidence["summary"]["path"]
        summary = check_perf_claims.load_json(summary_path)
        summary["results"][0]["elapsed_ns"]["legs_ns"]["a1"]["p50"] += 1
        summary_bytes = json.dumps(summary, sort_keys=True, indent=2).encode() + b"\n"
        summary_path.write_bytes(summary_bytes)
        manifest_path = package / self.evidence["manifest"]["path"]
        manifest = check_perf_claims.load_json(manifest_path)
        summary_identity = dict(manifest["summary_identity"])
        summary_identity.update(
            {
                "bytes": len(summary_bytes),
                "sha256": hashlib.sha256(summary_bytes).hexdigest(),
                "canonical_bytes": len(check_perf_claims.canonical_bytes(summary)),
                "canonical_sha256": check_perf_claims.canonical_sha256(summary),
            }
        )
        manifest["summary"] = summary_identity
        manifest["summary_identity"] = summary_identity
        manifest_path.write_text(
            json.dumps(manifest, sort_keys=True, indent=2) + "\n", encoding="utf-8"
        )
        self.evidence["summary"].update(
            {
                "sha256": summary_identity["sha256"],
                "canonical_sha256": summary_identity["canonical_sha256"],
            }
        )
        self.evidence["manifest"]["sha256"] = hashlib.sha256(
            manifest_path.read_bytes()
        ).hexdigest()
        with self.assertRaisesRegex(check_perf_claims.ClaimInputError, "differs from canonical"):
            self.verify()

    def test_declared_total_ceiling_is_checked_before_any_report_read(self) -> None:
        with patch.object(check_perf_claims, "_MAX_ABBA_DECOMPRESSED_BYTES", 1), patch.object(
            check_perf_claims, "_MAX_ABBA_MEMBER_BYTES", 1 << 60
        ):
            with self.assertRaisesRegex(check_perf_claims.ClaimInputError, "decompressed-byte ceiling"):
                self.verify()


if __name__ == "__main__":
    unittest.main()
