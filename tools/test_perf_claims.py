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
            {"abba-0248", "abba-0249", "abba-0250", "abba-0251", "resource-0251"},
        )
        self.assertEqual(
            set(claims),
            {
                "claim-0248-cfb-streaming",
                "claim-0249-ods-known-change",
                "claim-0250-zip-ordering",
                "claim-0251-xlsx-xml-borrowed",
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
