"""Unit tests for the CI smoke baseline selector, classifier and reporter.

Every branch of `tools/perf_smoke_baseline.py` is exercised here, because the
workflow itself cannot be run in this repository: a GitHub Actions run is the
only place the surrounding shell executes. The tests therefore cover the pure
decision logic end to end, drive the four CLI subcommands over real files, and
assert the workflow's wiring textually so the YAML cannot drift away from the
module it calls.
"""

from __future__ import annotations

import copy
import json
import re
import tempfile
import unittest
from pathlib import Path

from tools import perf_compare
from tools import perf_smoke_baseline
from tools.test_perf_compare import allocator_policy_fixture, allocator_report


ROOT = Path(__file__).resolve().parents[1]
WORKFLOW_PATH = ROOT / ".github" / "workflows" / "perf-baseline.yml"
SMOKE_POLICY_PATH = ROOT / "docs" / "performance" / "perf-smoke-baseline-policy-v1.json"
COMPARATOR_POLICY_PATH = (
    ROOT / "docs" / "performance" / "perf-regression-policy-allocator-v1.json"
)

RUNNER_LABELS = "ubuntu-latest"


def checked_smoke_policy() -> dict:
    return json.loads(SMOKE_POLICY_PATH.read_text(encoding="utf-8"))


def checked_comparator_policy() -> dict:
    return json.loads(COMPARATOR_POLICY_PATH.read_text(encoding="utf-8"))


def descriptor_for(
    report: dict,
    *,
    smoke_policy: dict,
    comparator_policy: dict,
    run_id: str = "4242",
    event: str = "schedule",
    runner_labels: str = RUNNER_LABELS,
) -> dict:
    return perf_smoke_baseline.build_descriptor(
        report,
        smoke_policy=smoke_policy,
        comparator_policy=comparator_policy,
        runner_labels=runner_labels,
        run_id=run_id,
        run_attempt=1,
        event=event,
    )


class SmokePolicyValidationTests(unittest.TestCase):
    def test_checked_policy_is_valid_and_advisory(self) -> None:
        policy = perf_smoke_baseline.validate_smoke_policy(checked_smoke_policy())
        self.assertEqual(policy["enforcement"], "advisory")
        self.assertIn("deliverable 8", policy["enforcement_reason"])
        self.assertTrue((ROOT / policy["comparator_policy"]).is_file())
        self.assertEqual(
            policy["comparator_policy_id"], checked_comparator_policy()["policy_id"]
        )

    def test_checked_policy_binds_the_comparator_policy_document(self) -> None:
        policy = checked_smoke_policy()
        self.assertEqual(
            Path(policy["comparator_policy"]),
            COMPARATOR_POLICY_PATH.relative_to(ROOT),
        )

    def test_missing_key_is_rejected(self) -> None:
        policy = checked_smoke_policy()
        del policy["enforcement"]
        with self.assertRaisesRegex(
            perf_smoke_baseline.SmokeBaselineError, "missing=.*enforcement"
        ):
            perf_smoke_baseline.validate_smoke_policy(policy)

    def test_unknown_key_is_rejected(self) -> None:
        policy = checked_smoke_policy()
        policy["surprise"] = True
        with self.assertRaisesRegex(
            perf_smoke_baseline.SmokeBaselineError, "unknown=.*surprise"
        ):
            perf_smoke_baseline.validate_smoke_policy(policy)

    def test_non_object_policy_is_rejected(self) -> None:
        with self.assertRaisesRegex(
            perf_smoke_baseline.SmokeBaselineError, "must be a JSON object"
        ):
            perf_smoke_baseline.validate_smoke_policy([])

    def test_schema_version_is_pinned(self) -> None:
        policy = checked_smoke_policy()
        policy["schema_version"] = 2
        with self.assertRaisesRegex(
            perf_smoke_baseline.SmokeBaselineError, "schema_version must be 1"
        ):
            perf_smoke_baseline.validate_smoke_policy(policy)

    def test_empty_text_field_is_rejected(self) -> None:
        policy = checked_smoke_policy()
        policy["reference_report_name"] = ""
        with self.assertRaisesRegex(
            perf_smoke_baseline.SmokeBaselineError, "reference_report_name"
        ):
            perf_smoke_baseline.validate_smoke_policy(policy)

    def test_reference_events_must_be_unique_non_empty(self) -> None:
        for events in ([], ["schedule", "schedule"], ["schedule", ""], "schedule"):
            policy = checked_smoke_policy()
            policy["reference_events"] = events
            with self.assertRaisesRegex(
                perf_smoke_baseline.SmokeBaselineError, "reference_events"
            ):
                perf_smoke_baseline.validate_smoke_policy(policy)

    def test_max_reference_runs_must_be_positive(self) -> None:
        for limit in (0, -1, True, "5"):
            policy = checked_smoke_policy()
            policy["max_reference_runs"] = limit
            with self.assertRaisesRegex(
                perf_smoke_baseline.SmokeBaselineError, "max_reference_runs"
            ):
                perf_smoke_baseline.validate_smoke_policy(policy)

    def test_enforcement_value_is_constrained(self) -> None:
        policy = checked_smoke_policy()
        policy["enforcement"] = "warn"
        with self.assertRaisesRegex(
            perf_smoke_baseline.SmokeBaselineError, "enforcement must be"
        ):
            perf_smoke_baseline.validate_smoke_policy(policy)

    def test_blocking_enforcement_is_accepted(self) -> None:
        policy = checked_smoke_policy()
        policy["enforcement"] = "blocking"
        self.assertEqual(
            perf_smoke_baseline.validate_smoke_policy(policy)["enforcement"], "blocking"
        )

    def test_require_flags_must_be_boolean(self) -> None:
        for field in ("require_runner_label_match", "require_result_key_manifest_match"):
            policy = checked_smoke_policy()
            policy[field] = "yes"
            with self.assertRaisesRegex(perf_smoke_baseline.SmokeBaselineError, field):
                perf_smoke_baseline.validate_smoke_policy(policy)

    def test_self_comparison_expectations_keys_are_exact(self) -> None:
        policy = checked_smoke_policy()
        del policy["self_comparison_expectations"]["compared_metrics"]
        with self.assertRaisesRegex(
            perf_smoke_baseline.SmokeBaselineError, "self_comparison_expectations"
        ):
            perf_smoke_baseline.validate_smoke_policy(policy)

    def test_self_comparison_expectations_types_are_checked(self) -> None:
        policy = checked_smoke_policy()
        policy["self_comparison_expectations"]["matched_results"] = -1
        with self.assertRaisesRegex(
            perf_smoke_baseline.SmokeBaselineError, "matched_results"
        ):
            perf_smoke_baseline.validate_smoke_policy(policy)
        policy = checked_smoke_policy()
        policy["self_comparison_expectations"]["status"] = 1
        with self.assertRaisesRegex(perf_smoke_baseline.SmokeBaselineError, "status"):
            perf_smoke_baseline.validate_smoke_policy(policy)
        policy = checked_smoke_policy()
        policy["self_comparison_expectations"] = []
        with self.assertRaisesRegex(
            perf_smoke_baseline.SmokeBaselineError, "self_comparison_expectations"
        ):
            perf_smoke_baseline.validate_smoke_policy(policy)


class RunnerLabelTests(unittest.TestCase):
    def test_comma_separated_labels_are_normalized(self) -> None:
        self.assertEqual(
            perf_smoke_baseline.normalize_runner_labels("ubuntu-latest, self-hosted"),
            ["self-hosted", "ubuntu-latest"],
        )

    def test_list_labels_are_deduplicated_and_sorted(self) -> None:
        self.assertEqual(
            perf_smoke_baseline.normalize_runner_labels(
                ["ubuntu-latest", "ubuntu-latest", " x64 "]
            ),
            ["ubuntu-latest", "x64"],
        )

    def test_empty_labels_are_rejected(self) -> None:
        for value in ("", " , ", []):
            with self.assertRaisesRegex(
                perf_smoke_baseline.SmokeBaselineError, "runner labels"
            ):
                perf_smoke_baseline.normalize_runner_labels(value)

    def test_non_string_labels_are_rejected(self) -> None:
        with self.assertRaisesRegex(
            perf_smoke_baseline.SmokeBaselineError, "runner labels must be strings"
        ):
            perf_smoke_baseline.normalize_runner_labels(["ubuntu-latest", 7])
        with self.assertRaisesRegex(
            perf_smoke_baseline.SmokeBaselineError, "must be a string or a list"
        ):
            perf_smoke_baseline.normalize_runner_labels(7)


class FetchStatusTests(unittest.TestCase):
    def test_a_successful_fetch_records_its_run(self) -> None:
        document = perf_smoke_baseline.fetch_status_document(
            fetched=True, run_id="4242", reason="downloaded"
        )
        self.assertEqual(
            document,
            {
                "schema_version": 1,
                "fetched": True,
                "reason": "downloaded",
                "run_id": "4242",
            },
        )

    def test_an_unsuccessful_fetch_records_only_its_reason(self) -> None:
        document = perf_smoke_baseline.fetch_status_document(
            fetched=False, run_id=None, reason="no prior run"
        )
        self.assertEqual(
            document,
            {"schema_version": 1, "fetched": False, "reason": "no prior run"},
        )

    def test_a_fetch_without_a_run_id_is_rejected(self) -> None:
        with self.assertRaisesRegex(
            perf_smoke_baseline.SmokeBaselineError, "must name its run id"
        ):
            perf_smoke_baseline.fetch_status_document(
                fetched=True, run_id=None, reason="downloaded"
            )

    def test_an_empty_reason_is_rejected(self) -> None:
        with self.assertRaisesRegex(
            perf_smoke_baseline.SmokeBaselineError, "fetch status reason"
        ):
            perf_smoke_baseline.fetch_status_document(
                fetched=False, run_id=None, reason=""
            )


class ChooseReferenceRunTests(unittest.TestCase):
    EVENTS = ("schedule", "workflow_dispatch")

    def choose(self, runs, *, current_sha="head", limit=10):
        return perf_smoke_baseline.choose_reference_runs(
            runs, current_sha=current_sha, events=self.EVENTS, limit=limit
        )

    def test_empty_listing_yields_no_candidate(self) -> None:
        self.assertEqual(self.choose([]), [])

    def test_non_list_listing_is_rejected(self) -> None:
        with self.assertRaisesRegex(
            perf_smoke_baseline.SmokeBaselineError, "must be a JSON array"
        ):
            self.choose({})

    def test_non_object_entry_is_rejected(self) -> None:
        with self.assertRaisesRegex(
            perf_smoke_baseline.SmokeBaselineError, "entry 0 must be an object"
        ):
            self.choose(["nope"])

    def test_unsuccessful_runs_are_skipped(self) -> None:
        runs = [
            {
                "databaseId": 1,
                "conclusion": "failure",
                "event": "schedule",
                "headSha": "a",
                "createdAt": "2026-09-01T00:00:00Z",
            }
        ]
        self.assertEqual(self.choose(runs), [])

    def test_other_events_are_skipped(self) -> None:
        runs = [
            {
                "databaseId": 1,
                "conclusion": "success",
                "event": "push",
                "headSha": "a",
                "createdAt": "2026-09-01T00:00:00Z",
            }
        ]
        self.assertEqual(self.choose(runs), [])

    def test_the_commit_under_test_is_skipped(self) -> None:
        runs = [
            {
                "databaseId": 1,
                "conclusion": "success",
                "event": "schedule",
                "headSha": "head",
                "createdAt": "2026-09-01T00:00:00Z",
            }
        ]
        self.assertEqual(self.choose(runs), [])
        self.assertEqual(len(self.choose(runs, current_sha="")), 1)

    def test_invalid_run_ids_are_skipped(self) -> None:
        runs = [
            {"databaseId": identifier, "conclusion": "success", "event": "schedule"}
            for identifier in ("5", 0, -1, True, None)
        ]
        self.assertEqual(self.choose(runs), [])

    def test_candidates_are_newest_first_and_limited(self) -> None:
        runs = [
            {
                "databaseId": index,
                "conclusion": "success",
                "event": "schedule",
                "headSha": f"sha{index}",
                "createdAt": f"2026-09-0{index}T00:00:00Z",
            }
            for index in (1, 3, 2)
        ]
        candidates = self.choose(runs)
        self.assertEqual([item["run_id"] for item in candidates], ["3", "2", "1"])
        self.assertEqual([item["run_id"] for item in self.choose(runs, limit=2)], ["3", "2"])

    def test_missing_optional_fields_are_tolerated(self) -> None:
        runs = [{"databaseId": 9, "conclusion": "success", "event": "workflow_dispatch"}]
        self.assertEqual(
            self.choose(runs),
            [
                {
                    "run_id": "9",
                    "event": "workflow_dispatch",
                    "head_sha": None,
                    "created_at": None,
                }
            ],
        )


class SelfComparisonBaselineTests(unittest.TestCase):
    def test_only_the_revision_label_changes(self) -> None:
        current = allocator_report(revision="deadbeef")
        baseline = perf_smoke_baseline.self_comparison_baseline(current)
        self.assertEqual(
            baseline["environment"]["git_revision"],
            "allocator-smoke-reference:deadbeef",
        )
        stripped_baseline = copy.deepcopy(baseline)
        stripped_current = copy.deepcopy(current)
        del stripped_baseline["environment"]["git_revision"]
        del stripped_current["environment"]["git_revision"]
        self.assertEqual(stripped_baseline, stripped_current)

    def test_the_current_report_is_not_mutated(self) -> None:
        current = allocator_report(revision="deadbeef")
        snapshot = copy.deepcopy(current)
        perf_smoke_baseline.self_comparison_baseline(current)
        self.assertEqual(current, snapshot)

    def test_a_missing_revision_is_rejected(self) -> None:
        current = allocator_report()
        del current["environment"]["git_revision"]
        with self.assertRaisesRegex(
            perf_smoke_baseline.SmokeBaselineError, "git_revision"
        ):
            perf_smoke_baseline.self_comparison_baseline(current)

    def test_a_non_object_report_is_rejected(self) -> None:
        with self.assertRaisesRegex(
            perf_smoke_baseline.SmokeBaselineError, "current report must be"
        ):
            perf_smoke_baseline.self_comparison_baseline("nope")

    def test_the_self_baseline_passes_the_comparator(self) -> None:
        comparator_policy = allocator_policy_fixture()
        current = allocator_report(revision="candidate")
        baseline = perf_smoke_baseline.self_comparison_baseline(current)
        comparison = perf_compare.compare_reports(baseline, current, comparator_policy)
        self.assertEqual(comparison["status"], "pass")
        self.assertEqual(comparison["summary"]["matched_results"], 2)
        self.assertEqual(comparison["summary"]["compared_metrics"], 20)


class DescriptorTests(unittest.TestCase):
    def setUp(self) -> None:
        self.smoke_policy = perf_smoke_baseline.validate_smoke_policy(
            checked_smoke_policy()
        )
        self.comparator_policy = allocator_policy_fixture()
        self.report = allocator_report(revision="reference")

    def test_descriptor_records_the_binding_fields(self) -> None:
        descriptor = descriptor_for(
            self.report,
            smoke_policy=self.smoke_policy,
            comparator_policy=self.comparator_policy,
        )
        self.assertEqual(descriptor["policy_id"], self.smoke_policy["policy_id"])
        self.assertEqual(
            descriptor["comparator_policy_id"], self.comparator_policy["policy_id"]
        )
        self.assertEqual(descriptor["git_revision"], "reference")
        self.assertEqual(descriptor["harness_binary_profile"], "release")
        self.assertEqual(descriptor["runner_labels"], ["ubuntu-latest"])
        self.assertEqual(
            descriptor["result_key_manifest_sha256"],
            self.comparator_policy["expected_result_keys_sha256"],
        )
        self.assertEqual(
            set(descriptor["environment"]),
            set(self.comparator_policy["build_identity_fields"]),
        )

    def test_a_debug_profile_report_is_refused(self) -> None:
        report = copy.deepcopy(self.report)
        report["binary_identity"]["profile"] = "debug"
        with self.assertRaisesRegex(
            perf_smoke_baseline.SmokeBaselineError, "binary profile"
        ):
            descriptor_for(
                report,
                smoke_policy=self.smoke_policy,
                comparator_policy=self.comparator_policy,
            )

    def test_a_report_without_binary_identity_is_refused(self) -> None:
        report = copy.deepcopy(self.report)
        del report["binary_identity"]
        with self.assertRaisesRegex(
            perf_smoke_baseline.SmokeBaselineError, "binary_identity"
        ):
            descriptor_for(
                report,
                smoke_policy=self.smoke_policy,
                comparator_policy=self.comparator_policy,
            )

    def test_a_report_without_environment_is_refused(self) -> None:
        report = copy.deepcopy(self.report)
        del report["environment"]
        with self.assertRaisesRegex(
            perf_smoke_baseline.SmokeBaselineError, "environment"
        ):
            descriptor_for(
                report,
                smoke_policy=self.smoke_policy,
                comparator_policy=self.comparator_policy,
            )

    def test_descriptor_validation_rejects_key_drift(self) -> None:
        descriptor = descriptor_for(
            self.report,
            smoke_policy=self.smoke_policy,
            comparator_policy=self.comparator_policy,
        )
        broken = copy.deepcopy(descriptor)
        del broken["run_id"]
        with self.assertRaisesRegex(
            perf_smoke_baseline.SmokeBaselineError, "missing=.*run_id"
        ):
            perf_smoke_baseline.validate_descriptor(broken)
        broken = copy.deepcopy(descriptor)
        broken["extra"] = 1
        with self.assertRaisesRegex(
            perf_smoke_baseline.SmokeBaselineError, "unknown=.*extra"
        ):
            perf_smoke_baseline.validate_descriptor(broken)

    def test_descriptor_validation_rejects_bad_scalars(self) -> None:
        descriptor = descriptor_for(
            self.report,
            smoke_policy=self.smoke_policy,
            comparator_policy=self.comparator_policy,
        )
        for field, value, pattern in (
            ("schema_version", 2, "schema_version must be 1"),
            ("run_attempt", 0, "run_attempt"),
            ("run_attempt", True, "run_attempt"),
            ("result_key_manifest_sha256", "A" * 64, "lowercase hex"),
            ("result_key_manifest_sha256", "ab", "lowercase hex"),
            ("tool", [], "descriptor.tool"),
            ("environment", [], "descriptor.environment"),
        ):
            broken = copy.deepcopy(descriptor)
            broken[field] = value
            with self.assertRaisesRegex(
                perf_smoke_baseline.SmokeBaselineError, pattern
            ):
                perf_smoke_baseline.validate_descriptor(broken)

    def test_descriptor_labels_are_normalized_in_place(self) -> None:
        descriptor = descriptor_for(
            self.report,
            smoke_policy=self.smoke_policy,
            comparator_policy=self.comparator_policy,
            runner_labels="ubuntu-latest,ubuntu-latest",
        )
        self.assertEqual(descriptor["runner_labels"], ["ubuntu-latest"])


class ReferenceCompatibilityTests(unittest.TestCase):
    def setUp(self) -> None:
        self.smoke_policy = perf_smoke_baseline.validate_smoke_policy(
            checked_smoke_policy()
        )
        self.comparator_policy = allocator_policy_fixture()
        self.current = allocator_report(revision="candidate")
        self.candidate = allocator_report(revision="reference")
        self.descriptor = descriptor_for(
            self.candidate,
            smoke_policy=self.smoke_policy,
            comparator_policy=self.comparator_policy,
        )

    def checks(self, **overrides) -> dict[str, dict]:
        arguments = {
            "current": self.current,
            "candidate": self.candidate,
            "descriptor": self.descriptor,
            "smoke_policy": self.smoke_policy,
            "comparator_policy": self.comparator_policy,
            "runner_labels": ["ubuntu-latest"],
            "fetched_run_id": "4242",
        }
        arguments.update(overrides)
        return {
            check["name"]: check
            for check in perf_smoke_baseline.check_reference(**arguments)
        }

    def assert_only_failure(self, checks: dict[str, dict], name: str) -> None:
        failures = sorted(key for key, check in checks.items() if not check["passed"])
        self.assertEqual(failures, [name])

    def test_a_compatible_reference_passes_every_check(self) -> None:
        checks = self.checks()
        self.assertTrue(all(check["passed"] for check in checks.values()))
        self.assertEqual(
            sorted(checks),
            [
                "corpus_key_manifest_sha256",
                "current_tool_identity",
                "descriptor_policy_identity",
                "descriptor_revision_identity",
                "descriptor_run_identity",
                "distinct_revisions",
                "harness_binary_profile",
                "reference_clean_worktree",
                "reference_descriptor",
                "reference_report_parsed",
                "reference_report_schema",
                "reference_tool_identity",
                "runner_labels",
            ],
        )

    def test_a_non_object_report_stops_at_the_first_check(self) -> None:
        checks = self.checks(candidate="not a report")
        self.assertEqual(list(checks), ["reference_report_parsed"])
        self.assertFalse(checks["reference_report_parsed"]["passed"])

    def test_a_wrong_report_schema_fails(self) -> None:
        candidate = copy.deepcopy(self.candidate)
        candidate["schema_version"] = 2
        self.assert_only_failure(
            self.checks(candidate=candidate), "reference_report_schema"
        )

    def test_a_reference_tool_identity_mismatch_fails(self) -> None:
        candidate = copy.deepcopy(self.candidate)
        candidate["tool"]["instrumentation"] = "none"
        checks = self.checks(candidate=candidate)
        self.assertFalse(checks["reference_tool_identity"]["passed"])

    def test_a_current_tool_identity_mismatch_fails(self) -> None:
        current = copy.deepcopy(self.current)
        current["tool"]["binary"] = "litchi-perf-baseline"
        self.assert_only_failure(self.checks(current=current), "current_tool_identity")

    def test_a_reference_binary_profile_mismatch_fails(self) -> None:
        candidate = copy.deepcopy(self.candidate)
        candidate["binary_identity"]["profile"] = "debug"
        self.assert_only_failure(
            self.checks(candidate=candidate), "harness_binary_profile"
        )

    def test_a_current_binary_profile_mismatch_fails(self) -> None:
        current = copy.deepcopy(self.current)
        del current["binary_identity"]
        self.assert_only_failure(self.checks(current=current), "harness_binary_profile")

    def test_a_missing_descriptor_fails_closed(self) -> None:
        checks = self.checks(descriptor=None)
        self.assertFalse(checks["reference_descriptor"]["passed"])
        self.assertNotIn("descriptor_policy_identity", checks)
        self.assertNotIn("runner_labels", checks)

    def test_a_descriptor_policy_mismatch_fails(self) -> None:
        descriptor = copy.deepcopy(self.descriptor)
        descriptor["comparator_policy_id"] = "some-other-policy"
        self.assert_only_failure(
            self.checks(descriptor=descriptor), "descriptor_policy_identity"
        )

    def test_a_runner_label_mismatch_fails(self) -> None:
        descriptor = copy.deepcopy(self.descriptor)
        descriptor["runner_labels"] = ["self-hosted"]
        self.assert_only_failure(self.checks(descriptor=descriptor), "runner_labels")

    def test_a_runner_label_mismatch_is_ignored_when_the_policy_allows_it(self) -> None:
        smoke_policy = copy.deepcopy(self.smoke_policy)
        smoke_policy["require_runner_label_match"] = False
        descriptor = copy.deepcopy(self.descriptor)
        descriptor["runner_labels"] = ["self-hosted"]
        checks = self.checks(descriptor=descriptor, smoke_policy=smoke_policy)
        self.assertNotIn("runner_labels", checks)
        self.assertTrue(all(check["passed"] for check in checks.values()))

    def test_a_run_identity_mismatch_fails(self) -> None:
        self.assert_only_failure(
            self.checks(fetched_run_id="9999"), "descriptor_run_identity"
        )

    def test_the_run_identity_check_is_skipped_without_a_run_id(self) -> None:
        checks = self.checks(fetched_run_id=None)
        self.assertNotIn("descriptor_run_identity", checks)

    def test_a_descriptor_revision_mismatch_fails(self) -> None:
        descriptor = copy.deepcopy(self.descriptor)
        descriptor["git_revision"] = "someone-elses-revision"
        self.assert_only_failure(
            self.checks(descriptor=descriptor), "descriptor_revision_identity"
        )

    def test_a_changed_corpus_fails_the_key_manifest_digest(self) -> None:
        candidate = copy.deepcopy(self.candidate)
        for result in candidate["results"]:
            result["corpus"]["archive_sha256"] = "changed"
        descriptor = descriptor_for(
            candidate,
            smoke_policy=self.smoke_policy,
            comparator_policy=self.comparator_policy,
        )
        checks = self.checks(candidate=candidate, descriptor=descriptor)
        self.assertFalse(checks["corpus_key_manifest_sha256"]["passed"])

    def test_an_uncomputable_key_manifest_digest_fails(self) -> None:
        candidate = copy.deepcopy(self.candidate)
        candidate["results"] = candidate["results"][:1]
        checks = self.checks(candidate=candidate)
        self.assertFalse(checks["corpus_key_manifest_sha256"]["passed"])
        self.assertIn("could not be computed", checks["corpus_key_manifest_sha256"]["detail"])

    def test_a_forged_descriptor_digest_fails(self) -> None:
        descriptor = copy.deepcopy(self.descriptor)
        descriptor["result_key_manifest_sha256"] = "0" * 64
        checks = self.checks(descriptor=descriptor)
        self.assertFalse(checks["corpus_key_manifest_sha256"]["passed"])
        self.assertIn("descriptor digest disagrees", checks["corpus_key_manifest_sha256"]["detail"])

    def test_the_digest_check_is_skipped_when_the_policy_allows_it(self) -> None:
        smoke_policy = copy.deepcopy(self.smoke_policy)
        smoke_policy["require_result_key_manifest_match"] = False
        checks = self.checks(smoke_policy=smoke_policy)
        self.assertNotIn("corpus_key_manifest_sha256", checks)

    def test_an_identical_revision_fails(self) -> None:
        candidate = copy.deepcopy(self.candidate)
        candidate["environment"]["git_revision"] = self.current["environment"][
            "git_revision"
        ]
        descriptor = descriptor_for(
            candidate,
            smoke_policy=self.smoke_policy,
            comparator_policy=self.comparator_policy,
        )
        self.assert_only_failure(
            self.checks(candidate=candidate, descriptor=descriptor),
            "distinct_revisions",
        )

    def test_a_dirty_reference_worktree_fails(self) -> None:
        candidate = copy.deepcopy(self.candidate)
        candidate["environment"]["git_worktree_dirty"] = True
        self.assert_only_failure(
            self.checks(candidate=candidate), "reference_clean_worktree"
        )


class SelectBaselineTests(unittest.TestCase):
    def setUp(self) -> None:
        self.smoke_policy = perf_smoke_baseline.validate_smoke_policy(
            checked_smoke_policy()
        )
        self.comparator_policy = allocator_policy_fixture()
        self.current = allocator_report(revision="candidate")
        self.candidate = allocator_report(revision="reference")
        self.descriptor = descriptor_for(
            self.candidate,
            smoke_policy=self.smoke_policy,
            comparator_policy=self.comparator_policy,
        )

    def select(self, **overrides):
        arguments = {
            "current": self.current,
            "candidate": self.candidate,
            "descriptor": self.descriptor,
            "fetch_status": {"fetched": True, "run_id": "4242"},
            "smoke_policy": self.smoke_policy,
            "comparator_policy": self.comparator_policy,
            "runner_labels": RUNNER_LABELS,
            "current_run_id": "9001",
        }
        arguments.update(overrides)
        return perf_smoke_baseline.select_baseline(**arguments)

    def test_a_compatible_reference_is_used(self) -> None:
        baseline, selection = self.select()
        self.assertEqual(selection["mode"], perf_smoke_baseline.MODE_FETCHED)
        self.assertTrue(selection["is_regression_detector"])
        self.assertEqual(selection["fallback_reasons"], [])
        self.assertEqual(baseline, self.candidate)
        self.assertEqual(selection["reference"]["run_id"], "4242")
        self.assertEqual(selection["reference"]["event"], "schedule")
        self.assertEqual(selection["reference"]["git_revision"], "reference")
        self.assertEqual(selection["current"]["run_id"], "9001")
        self.assertIn("4242", selection["label"])

    def test_no_fetched_artifact_falls_back_with_the_recorded_reason(self) -> None:
        baseline, selection = self.select(
            fetch_status={"fetched": False, "reason": "no prior full run"}
        )
        self.assertEqual(selection["mode"], perf_smoke_baseline.MODE_SELF)
        self.assertFalse(selection["is_regression_detector"])
        self.assertEqual(
            selection["fallback_reasons"], ["reference_fetched: no prior full run"]
        )
        self.assertIn("plumbing check", selection["label"])
        self.assertEqual(
            baseline["environment"]["git_revision"],
            "allocator-smoke-reference:candidate",
        )

    def test_a_missing_fetch_reason_still_falls_back(self) -> None:
        _, selection = self.select(fetch_status={"fetched": False})
        self.assertEqual(selection["mode"], perf_smoke_baseline.MODE_SELF)
        self.assertEqual(
            selection["fallback_reasons"],
            ["reference_fetched: no reference artifact was fetched"],
        )

    def test_an_absent_fetch_status_falls_back(self) -> None:
        _, selection = self.select(fetch_status=None)
        self.assertEqual(selection["mode"], perf_smoke_baseline.MODE_SELF)

    def test_an_incompatible_reference_falls_back_and_lists_every_reason(self) -> None:
        candidate = copy.deepcopy(self.candidate)
        candidate["binary_identity"]["profile"] = "debug"
        candidate["environment"]["git_worktree_dirty"] = True
        _, selection = self.select(candidate=candidate)
        self.assertEqual(selection["mode"], perf_smoke_baseline.MODE_SELF)
        self.assertEqual(len(selection["fallback_reasons"]), 2)
        self.assertTrue(
            any("harness_binary_profile" in reason for reason in selection["fallback_reasons"])
        )
        self.assertTrue(
            any(
                "reference_clean_worktree" in reason
                for reason in selection["fallback_reasons"]
            )
        )

    def test_the_selected_reference_baseline_is_accepted_by_the_comparator(self) -> None:
        baseline, selection = self.select()
        comparison = perf_compare.compare_reports(
            baseline, self.current, self.comparator_policy
        )
        self.assertEqual(selection["mode"], perf_smoke_baseline.MODE_FETCHED)
        self.assertEqual(comparison["status"], "pass")

    def test_a_reference_regression_is_visible_through_the_comparator(self) -> None:
        candidate = allocator_report(value=100, revision="reference")
        descriptor = descriptor_for(
            candidate,
            smoke_policy=self.smoke_policy,
            comparator_policy=self.comparator_policy,
        )
        current = allocator_report(value=200, revision="candidate")
        baseline, selection = self.select(
            candidate=candidate, descriptor=descriptor, current=current
        )
        self.assertEqual(selection["mode"], perf_smoke_baseline.MODE_FETCHED)
        comparison = perf_compare.compare_reports(
            baseline, current, self.comparator_policy
        )
        self.assertEqual(comparison["status"], "regression")
        self.assertGreater(comparison["summary"]["regressions"], 0)

    def test_a_non_object_current_report_is_rejected(self) -> None:
        with self.assertRaisesRegex(
            perf_smoke_baseline.SmokeBaselineError, "current report"
        ):
            self.select(current=[])

    def test_a_non_object_fetch_status_is_rejected(self) -> None:
        with self.assertRaisesRegex(
            perf_smoke_baseline.SmokeBaselineError, "fetch status"
        ):
            self.select(fetch_status=[])


class ClassificationTests(unittest.TestCase):
    def setUp(self) -> None:
        self.smoke_policy = perf_smoke_baseline.validate_smoke_policy(
            checked_smoke_policy()
        )
        self.blocking_policy = copy.deepcopy(self.smoke_policy)
        self.blocking_policy["enforcement"] = "blocking"
        self.comparator_policy = allocator_policy_fixture()

    def selection(self, mode):
        return {
            "schema_version": perf_smoke_baseline.SELECTION_SCHEMA,
            "mode": mode,
            "label": "label",
            "is_regression_detector": mode == perf_smoke_baseline.MODE_FETCHED,
            "reference": {"run_id": "4242", "git_revision": "reference"}
            if mode == perf_smoke_baseline.MODE_FETCHED
            else None,
            "fallback_reasons": []
            if mode == perf_smoke_baseline.MODE_FETCHED
            else ["reference_fetched: none"],
        }

    def self_comparison(self):
        current = allocator_report(revision="candidate")
        baseline = perf_smoke_baseline.self_comparison_baseline(current)
        return perf_compare.compare_reports(baseline, current, self.comparator_policy)

    def classify(self, comparison, *, mode, exit_status=0, policy=None):
        return perf_smoke_baseline.classify_comparison(
            comparison=comparison,
            comparator_exit_status=exit_status,
            selection=self.selection(mode),
            smoke_policy=policy or self.smoke_policy,
        )

    def test_a_healthy_plumbing_check_passes(self) -> None:
        classification = self.classify(
            self.self_comparison(), mode=perf_smoke_baseline.MODE_SELF
        )
        self.assertEqual(
            classification["outcome"], perf_smoke_baseline.OUTCOME_PLUMBING_PASS
        )
        self.assertEqual(classification["exit_status"], 0)
        self.assertFalse(classification["blocking"])

    def test_the_checked_expectations_match_the_fixture_shape(self) -> None:
        comparison = self.self_comparison()
        expectations = self.smoke_policy["self_comparison_expectations"]
        self.assertEqual(comparison["status"], expectations["status"])
        for field, expected in expectations.items():
            if field == "status":
                continue
            self.assertEqual(comparison["summary"][field], expected, field)

    def test_a_wrong_summary_number_is_a_plumbing_defect(self) -> None:
        comparison = self.self_comparison()
        comparison["summary"]["compared_metrics"] = 19
        classification = self.classify(comparison, mode=perf_smoke_baseline.MODE_SELF)
        self.assertEqual(
            classification["outcome"], perf_smoke_baseline.OUTCOME_PLUMBING_DEFECT
        )
        self.assertEqual(classification["exit_status"], 2)
        self.assertTrue(classification["blocking"])

    def test_a_missing_summary_is_a_plumbing_defect(self) -> None:
        comparison = self.self_comparison()
        del comparison["summary"]
        classification = self.classify(comparison, mode=perf_smoke_baseline.MODE_SELF)
        self.assertEqual(
            classification["outcome"], perf_smoke_baseline.OUTCOME_PLUMBING_DEFECT
        )
        self.assertIn("comparison has no summary object", classification["detail"])

    def test_an_invalid_self_comparison_is_a_plumbing_defect(self) -> None:
        comparison = perf_compare.invalid_report(ValueError("policy drifted"))
        classification = self.classify(
            comparison, mode=perf_smoke_baseline.MODE_SELF, exit_status=2
        )
        self.assertEqual(
            classification["outcome"], perf_smoke_baseline.OUTCOME_PLUMBING_DEFECT
        )
        self.assertEqual(classification["exit_status"], 2)
        self.assertTrue(
            any("policy drifted" in item for item in classification["detail"])
        )

    def test_a_nonzero_comparator_status_is_a_plumbing_defect(self) -> None:
        classification = self.classify(
            self.self_comparison(), mode=perf_smoke_baseline.MODE_SELF, exit_status=1
        )
        self.assertEqual(
            classification["outcome"], perf_smoke_baseline.OUTCOME_PLUMBING_DEFECT
        )
        self.assertIn("comparator exit status 1", classification["detail"])

    def test_a_plumbing_defect_blocks_even_under_advisory_enforcement(self) -> None:
        comparison = self.self_comparison()
        comparison["status"] = "regression"
        classification = self.classify(comparison, mode=perf_smoke_baseline.MODE_SELF)
        self.assertEqual(classification["enforcement"], "advisory")
        self.assertTrue(classification["blocking"])
        self.assertEqual(classification["exit_status"], 2)

    def test_a_clean_reference_comparison_passes(self) -> None:
        comparison = self.self_comparison()
        classification = self.classify(comparison, mode=perf_smoke_baseline.MODE_FETCHED)
        self.assertEqual(
            classification["outcome"], perf_smoke_baseline.OUTCOME_REFERENCE_PASS
        )
        self.assertEqual(classification["exit_status"], 0)

    def test_a_reference_regression_is_advisory_by_default(self) -> None:
        comparison = self.self_comparison()
        comparison["status"] = "regression"
        comparison["regressions"] = [
            {
                "case": "opc_file_eager_open",
                "metric": "operation_metrics/allocation/allocation_calls/values",
                "baseline": 100,
                "current": 200,
            },
            "not an object",
        ]
        classification = self.classify(comparison, mode=perf_smoke_baseline.MODE_FETCHED)
        self.assertEqual(
            classification["outcome"], perf_smoke_baseline.OUTCOME_REFERENCE_REGRESSION
        )
        self.assertFalse(classification["blocking"])
        self.assertEqual(classification["exit_status"], 0)
        self.assertEqual(len(classification["detail"]), 1)

    def test_a_reference_regression_blocks_when_the_policy_says_so(self) -> None:
        comparison = self.self_comparison()
        comparison["status"] = "regression"
        classification = self.classify(
            comparison,
            mode=perf_smoke_baseline.MODE_FETCHED,
            policy=self.blocking_policy,
        )
        self.assertTrue(classification["blocking"])
        self.assertEqual(classification["exit_status"], 1)

    def test_build_identity_drift_is_reported_and_never_blocks(self) -> None:
        comparison = perf_compare.invalid_report(
            ValueError("build identity mismatch for 'cpu_model': 'A' != 'B'")
        )
        for policy in (self.smoke_policy, self.blocking_policy):
            classification = self.classify(
                comparison,
                mode=perf_smoke_baseline.MODE_FETCHED,
                exit_status=2,
                policy=policy,
            )
            self.assertEqual(
                classification["outcome"],
                perf_smoke_baseline.OUTCOME_REFERENCE_ENVIRONMENT_DRIFT,
            )
            self.assertFalse(classification["blocking"])
            self.assertEqual(classification["exit_status"], 0)

    def test_a_revision_collision_counts_as_drift(self) -> None:
        comparison = perf_compare.invalid_report(
            ValueError("reference and current git revisions must differ")
        )
        classification = self.classify(
            comparison, mode=perf_smoke_baseline.MODE_FETCHED, exit_status=2
        )
        self.assertEqual(
            classification["outcome"],
            perf_smoke_baseline.OUTCOME_REFERENCE_ENVIRONMENT_DRIFT,
        )

    def test_a_mixed_error_set_is_an_input_defect(self) -> None:
        comparison = perf_compare.invalid_report(ValueError("x"))
        comparison["errors"] = [
            "build identity mismatch for 'cpu_model': 'A' != 'B'",
            "case/corpus manifest digest does not match policy",
        ]
        classification = self.classify(
            comparison, mode=perf_smoke_baseline.MODE_FETCHED, exit_status=2
        )
        self.assertEqual(
            classification["outcome"],
            perf_smoke_baseline.OUTCOME_REFERENCE_INPUT_DEFECT,
        )
        self.assertFalse(classification["blocking"])
        self.assertEqual(classification["exit_status"], 0)

    def test_an_input_defect_blocks_when_the_policy_says_so(self) -> None:
        comparison = perf_compare.invalid_report(
            ValueError("case/corpus manifest digest does not match policy")
        )
        classification = self.classify(
            comparison,
            mode=perf_smoke_baseline.MODE_FETCHED,
            exit_status=2,
            policy=self.blocking_policy,
        )
        self.assertEqual(
            classification["outcome"],
            perf_smoke_baseline.OUTCOME_REFERENCE_INPUT_DEFECT,
        )
        self.assertTrue(classification["blocking"])
        self.assertEqual(classification["exit_status"], 2)

    def test_an_invalid_report_without_errors_is_an_input_defect(self) -> None:
        comparison = perf_compare.invalid_report(ValueError("x"))
        comparison["errors"] = []
        classification = self.classify(
            comparison, mode=perf_smoke_baseline.MODE_FETCHED, exit_status=2
        )
        self.assertEqual(
            classification["outcome"],
            perf_smoke_baseline.OUTCOME_REFERENCE_INPUT_DEFECT,
        )

    def test_an_unknown_comparator_status_is_rejected(self) -> None:
        comparison = self.self_comparison()
        comparison["status"] = "confused"
        with self.assertRaisesRegex(
            perf_smoke_baseline.SmokeBaselineError, "status 'confused' is unknown"
        ):
            self.classify(comparison, mode=perf_smoke_baseline.MODE_FETCHED)

    def test_a_non_object_comparison_is_rejected(self) -> None:
        with self.assertRaisesRegex(
            perf_smoke_baseline.SmokeBaselineError, "comparison must be"
        ):
            self.classify([], mode=perf_smoke_baseline.MODE_FETCHED)

    def test_selection_validation_rejects_unknown_modes(self) -> None:
        with self.assertRaisesRegex(
            perf_smoke_baseline.SmokeBaselineError, "mode 'mystery' is unknown"
        ):
            perf_smoke_baseline.validate_selection(
                {"schema_version": 1, "mode": "mystery"}
            )
        with self.assertRaisesRegex(
            perf_smoke_baseline.SmokeBaselineError, "schema_version must be 1"
        ):
            perf_smoke_baseline.validate_selection({"schema_version": 9, "mode": "x"})


class RenderingTests(unittest.TestCase):
    setUp = ClassificationTests.setUp
    selection = ClassificationTests.selection
    self_comparison = ClassificationTests.self_comparison
    classify = ClassificationTests.classify

    def test_the_fallback_summary_says_it_detects_nothing(self) -> None:
        selection = self.selection(perf_smoke_baseline.MODE_SELF)
        selection["label"] = "plumbing check only"
        classification = self.classify(
            self.self_comparison(), mode=perf_smoke_baseline.MODE_SELF
        )
        summary = perf_smoke_baseline.render_summary(selection, classification)
        self.assertIn("Detects regressions: **no**", summary)
        self.assertIn("plumbing check only", summary)
        self.assertIn("Why the fetched reference was not used", summary)

    def test_the_reference_summary_names_the_run(self) -> None:
        selection = self.selection(perf_smoke_baseline.MODE_FETCHED)
        classification = self.classify(
            self.self_comparison(), mode=perf_smoke_baseline.MODE_FETCHED
        )
        summary = perf_smoke_baseline.render_summary(selection, classification)
        self.assertIn("Detects regressions: yes", summary)
        self.assertIn("run `4242`", summary)
        self.assertNotIn("Why the fetched reference was not used", summary)

    def test_annotation_levels_follow_the_outcome(self) -> None:
        cases = {
            perf_smoke_baseline.OUTCOME_PLUMBING_PASS: "::notice",
            perf_smoke_baseline.OUTCOME_REFERENCE_PASS: "::notice",
            perf_smoke_baseline.OUTCOME_REFERENCE_ENVIRONMENT_DRIFT: "::notice",
            perf_smoke_baseline.OUTCOME_REFERENCE_REGRESSION: "::warning",
            perf_smoke_baseline.OUTCOME_REFERENCE_INPUT_DEFECT: "::warning",
        }
        for outcome, prefix in cases.items():
            mode = (
                perf_smoke_baseline.MODE_SELF
                if outcome == perf_smoke_baseline.OUTCOME_PLUMBING_PASS
                else perf_smoke_baseline.MODE_FETCHED
            )
            selection = self.selection(mode)
            classification = {
                "outcome": outcome,
                "blocking": False,
                "detail": [],
                "enforcement": "advisory",
                "comparator_status": "pass",
            }
            annotation = perf_smoke_baseline.render_annotations(
                selection, classification
            )[0]
            self.assertTrue(annotation.startswith(prefix), annotation)

    def test_a_blocking_outcome_annotates_as_an_error(self) -> None:
        selection = self.selection(perf_smoke_baseline.MODE_SELF)
        classification = {
            "outcome": perf_smoke_baseline.OUTCOME_PLUMBING_DEFECT,
            "blocking": True,
            "detail": ["summary.compared_metrics is 19, expected 20"],
            "enforcement": "advisory",
            "comparator_status": "pass",
        }
        annotation = perf_smoke_baseline.render_annotations(selection, classification)[0]
        self.assertTrue(annotation.startswith("::error"))
        self.assertIn("fallback because", annotation)

    def test_annotations_escape_newlines_and_percent_signs(self) -> None:
        selection = self.selection(perf_smoke_baseline.MODE_FETCHED)
        classification = {
            "outcome": perf_smoke_baseline.OUTCOME_REFERENCE_REGRESSION,
            "blocking": False,
            "detail": ["allocation_calls +5%\nsecond line"],
            "enforcement": "advisory",
            "comparator_status": "regression",
        }
        annotation = perf_smoke_baseline.render_annotations(selection, classification)[0]
        self.assertNotIn("\n", annotation)
        self.assertIn("%25", annotation)
        self.assertIn("%0A", annotation)
        self.assertIn("advisory: reported, not enforced", annotation)


class CommandLineTests(unittest.TestCase):
    def setUp(self) -> None:
        self.directory = tempfile.TemporaryDirectory()
        self.addCleanup(self.directory.cleanup)
        self.root = Path(self.directory.name)
        self.smoke_policy = perf_smoke_baseline.validate_smoke_policy(
            checked_smoke_policy()
        )
        self.comparator_policy = allocator_policy_fixture()
        self.smoke_policy_path = self.root / "smoke-policy.json"
        self.comparator_policy_path = self.root / "comparator-policy.json"
        self._write(self.smoke_policy_path, checked_smoke_policy())
        self._write(self.comparator_policy_path, self.comparator_policy)
        self.current_path = self.root / "current.json"
        self._write(self.current_path, allocator_report(revision="candidate"))
        self.reference_dir = self.root / "reference"
        self.reference_dir.mkdir()

    def _write(self, path: Path, document) -> None:
        path.write_text(json.dumps(document, indent=2, sort_keys=True), encoding="utf-8")

    def _read(self, path: Path):
        return json.loads(path.read_text(encoding="utf-8"))

    def publish_reference(self, *, run_id="4242", revision="reference") -> None:
        report = allocator_report(revision=revision)
        report_path = self.root / "full-run.json"
        self._write(report_path, report)
        descriptor_path = self.reference_dir / self.smoke_policy[
            "reference_descriptor_name"
        ]
        status = perf_smoke_baseline.main(
            [
                "descriptor",
                "--policy",
                str(self.smoke_policy_path),
                "--comparator-policy",
                str(self.comparator_policy_path),
                "--report",
                str(report_path),
                "--runner-labels",
                RUNNER_LABELS,
                "--run-id",
                run_id,
                "--event",
                "schedule",
                "--out",
                str(descriptor_path),
            ]
        )
        self.assertEqual(status, 0)
        self._write(
            self.reference_dir / self.smoke_policy["reference_report_name"], report
        )

    def select(self, *, fetch_status) -> tuple[Path, Path]:
        status_path = self.root / "fetch-status.json"
        self._write(status_path, fetch_status)
        baseline_path = self.root / "baseline.json"
        selection_path = self.root / "selection.json"
        code = perf_smoke_baseline.main(
            [
                "select",
                "--policy",
                str(self.smoke_policy_path),
                "--comparator-policy",
                str(self.comparator_policy_path),
                "--current",
                str(self.current_path),
                "--reference-dir",
                str(self.reference_dir),
                "--fetch-status",
                str(status_path),
                "--runner-labels",
                RUNNER_LABELS,
                "--current-run-id",
                "9001",
                "--baseline-out",
                str(baseline_path),
                "--selection-out",
                str(selection_path),
            ]
        )
        self.assertEqual(code, 0)
        return baseline_path, selection_path

    def compare(self, baseline_path: Path) -> tuple[Path, int]:
        comparison_path = self.root / "comparison.json"
        code = perf_compare.main(
            [
                "--policy",
                str(self.comparator_policy_path),
                "--baseline",
                str(baseline_path),
                "--current",
                str(self.current_path),
                "--json-out",
                str(comparison_path),
            ]
        )
        return comparison_path, code

    def report(self, selection_path: Path, comparison_path: Path, exit_status: int) -> tuple[int, dict, str]:
        classification_path = self.root / "classification.json"
        summary_path = self.root / "outcome.md"
        step_summary_path = self.root / "step-summary.md"
        code = perf_smoke_baseline.main(
            [
                "report",
                "--policy",
                str(self.smoke_policy_path),
                "--selection",
                str(selection_path),
                "--comparison",
                str(comparison_path),
                "--comparator-exit-status",
                str(exit_status),
                "--classification-out",
                str(classification_path),
                "--summary-out",
                str(summary_path),
                "--step-summary",
                str(step_summary_path),
            ]
        )
        return (
            code,
            self._read(classification_path),
            step_summary_path.read_text(encoding="utf-8"),
        )

    def test_the_fallback_path_runs_end_to_end(self) -> None:
        baseline_path, selection_path = self.select(
            fetch_status={"fetched": False, "reason": "no prior successful full run"}
        )
        selection = self._read(selection_path)
        self.assertEqual(selection["mode"], perf_smoke_baseline.MODE_SELF)
        comparison_path, comparator_status = self.compare(baseline_path)
        self.assertEqual(comparator_status, 0)
        code, classification, step_summary = self.report(
            selection_path, comparison_path, comparator_status
        )
        self.assertEqual(code, 0)
        self.assertEqual(
            classification["outcome"], perf_smoke_baseline.OUTCOME_PLUMBING_PASS
        )
        self.assertIn("Detects regressions: **no**", step_summary)

    def test_the_fetched_path_runs_end_to_end(self) -> None:
        self.publish_reference()
        baseline_path, selection_path = self.select(
            fetch_status={"fetched": True, "run_id": "4242"}
        )
        selection = self._read(selection_path)
        self.assertEqual(selection["mode"], perf_smoke_baseline.MODE_FETCHED)
        comparison_path, comparator_status = self.compare(baseline_path)
        self.assertEqual(comparator_status, 0)
        code, classification, step_summary = self.report(
            selection_path, comparison_path, comparator_status
        )
        self.assertEqual(code, 0)
        self.assertEqual(
            classification["outcome"], perf_smoke_baseline.OUTCOME_REFERENCE_PASS
        )
        self.assertIn("run `4242`", step_summary)

    def test_a_fetched_regression_is_reported_without_failing(self) -> None:
        self.publish_reference()
        self._write(self.current_path, allocator_report(value=200, revision="candidate"))
        baseline_path, selection_path = self.select(
            fetch_status={"fetched": True, "run_id": "4242"}
        )
        comparison_path, comparator_status = self.compare(baseline_path)
        self.assertEqual(comparator_status, 1)
        code, classification, _ = self.report(
            selection_path, comparison_path, comparator_status
        )
        self.assertEqual(
            classification["outcome"],
            perf_smoke_baseline.OUTCOME_REFERENCE_REGRESSION,
        )
        self.assertEqual(code, 0)
        self.assertFalse(classification["blocking"])

    def test_an_artifact_without_the_report_falls_back(self) -> None:
        _, selection_path = self.select(fetch_status={"fetched": True, "run_id": "4242"})
        selection = self._read(selection_path)
        self.assertEqual(selection["mode"], perf_smoke_baseline.MODE_SELF)
        self.assertIn(
            "carries no", " ".join(selection["fallback_reasons"])
        )

    def test_an_unreadable_reference_report_falls_back(self) -> None:
        self.publish_reference()
        (self.reference_dir / self.smoke_policy["reference_report_name"]).write_text(
            "{ not json", encoding="utf-8"
        )
        _, selection_path = self.select(fetch_status={"fetched": True, "run_id": "4242"})
        selection = self._read(selection_path)
        self.assertEqual(selection["mode"], perf_smoke_baseline.MODE_SELF)

    def test_an_unreadable_descriptor_falls_back(self) -> None:
        self.publish_reference()
        (self.reference_dir / self.smoke_policy["reference_descriptor_name"]).write_text(
            "{ not json", encoding="utf-8"
        )
        _, selection_path = self.select(fetch_status={"fetched": True, "run_id": "4242"})
        selection = self._read(selection_path)
        self.assertEqual(selection["mode"], perf_smoke_baseline.MODE_SELF)

    def test_an_absent_reference_directory_falls_back(self) -> None:
        status_path = self.root / "fetch-status.json"
        self._write(status_path, {"fetched": False, "reason": "gh CLI unavailable"})
        selection_path = self.root / "selection.json"
        code = perf_smoke_baseline.main(
            [
                "select",
                "--policy",
                str(self.smoke_policy_path),
                "--comparator-policy",
                str(self.comparator_policy_path),
                "--current",
                str(self.current_path),
                "--fetch-status",
                str(status_path),
                "--runner-labels",
                RUNNER_LABELS,
                "--baseline-out",
                str(self.root / "baseline.json"),
                "--selection-out",
                str(selection_path),
            ]
        )
        self.assertEqual(code, 0)
        self.assertEqual(
            self._read(selection_path)["mode"], perf_smoke_baseline.MODE_SELF
        )

    def test_a_missing_fetch_status_file_falls_back(self) -> None:
        selection_path = self.root / "selection.json"
        code = perf_smoke_baseline.main(
            [
                "select",
                "--policy",
                str(self.smoke_policy_path),
                "--comparator-policy",
                str(self.comparator_policy_path),
                "--current",
                str(self.current_path),
                "--fetch-status",
                str(self.root / "absent.json"),
                "--runner-labels",
                RUNNER_LABELS,
                "--baseline-out",
                str(self.root / "baseline.json"),
                "--selection-out",
                str(selection_path),
            ]
        )
        self.assertEqual(code, 0)
        self.assertEqual(
            self._read(selection_path)["mode"], perf_smoke_baseline.MODE_SELF
        )

    def test_choose_runs_writes_one_identifier_per_line(self) -> None:
        runs_path = self.root / "runs.json"
        self._write(
            runs_path,
            [
                {
                    "databaseId": 11,
                    "conclusion": "success",
                    "event": "schedule",
                    "headSha": "old",
                    "createdAt": "2026-09-01T00:00:00Z",
                },
                {
                    "databaseId": 12,
                    "conclusion": "success",
                    "event": "push",
                    "headSha": "other",
                    "createdAt": "2026-09-02T00:00:00Z",
                },
                {
                    "databaseId": 13,
                    "conclusion": "success",
                    "event": "workflow_dispatch",
                    "headSha": "head",
                    "createdAt": "2026-09-03T00:00:00Z",
                },
            ],
        )
        out = self.root / "candidates.txt"
        code = perf_smoke_baseline.main(
            [
                "choose-runs",
                "--policy",
                str(self.smoke_policy_path),
                "--runs",
                str(runs_path),
                "--current-sha",
                "head",
                "--out",
                str(out),
            ]
        )
        self.assertEqual(code, 0)
        self.assertEqual(out.read_text(encoding="utf-8"), "11\n")

    def test_fetch_status_writes_both_shapes(self) -> None:
        status_path = self.root / "fetch-status.json"
        self.assertEqual(
            perf_smoke_baseline.main(
                [
                    "fetch-status",
                    "--out",
                    str(status_path),
                    "--reason",
                    "no prior successful full run",
                ]
            ),
            0,
        )
        self.assertEqual(
            self._read(status_path),
            {
                "schema_version": 1,
                "fetched": False,
                "reason": "no prior successful full run",
            },
        )
        self.assertEqual(
            perf_smoke_baseline.main(
                [
                    "fetch-status",
                    "--out",
                    str(status_path),
                    "--fetched",
                    "--run-id",
                    "4242",
                    "--reason",
                    "downloaded the baseline artifact of run 4242",
                ]
            ),
            0,
        )
        self.assertEqual(self._read(status_path)["run_id"], "4242")

    def test_fetch_status_refuses_a_fetch_without_a_run(self) -> None:
        self.assertEqual(
            perf_smoke_baseline.main(
                [
                    "fetch-status",
                    "--out",
                    str(self.root / "fetch-status.json"),
                    "--fetched",
                    "--reason",
                    "downloaded",
                ]
            ),
            2,
        )

    def test_the_fetch_status_the_workflow_writes_drives_the_selector(self) -> None:
        status_path = self.root / "fetch-status.json"
        perf_smoke_baseline.main(
            [
                "fetch-status",
                "--out",
                str(status_path),
                "--reason",
                "the gh CLI is not available on this runner",
            ]
        )
        _, selection_path = self.select(fetch_status=self._read(status_path))
        self.assertEqual(
            self._read(selection_path)["fallback_reasons"],
            ["reference_fetched: the gh CLI is not available on this runner"],
        )

    def test_an_unreadable_input_exits_two(self) -> None:
        code = perf_smoke_baseline.main(
            [
                "choose-runs",
                "--policy",
                str(self.smoke_policy_path),
                "--runs",
                str(self.root / "absent.json"),
                "--out",
                str(self.root / "candidates.txt"),
            ]
        )
        self.assertEqual(code, 2)

    def test_a_malformed_current_report_exits_two(self) -> None:
        self.current_path.write_text("{ not json", encoding="utf-8")
        status_path = self.root / "fetch-status.json"
        self._write(status_path, {"fetched": False, "reason": "none"})
        code = perf_smoke_baseline.main(
            [
                "select",
                "--policy",
                str(self.smoke_policy_path),
                "--comparator-policy",
                str(self.comparator_policy_path),
                "--current",
                str(self.current_path),
                "--fetch-status",
                str(status_path),
                "--runner-labels",
                RUNNER_LABELS,
                "--baseline-out",
                str(self.root / "baseline.json"),
                "--selection-out",
                str(self.root / "selection.json"),
            ]
        )
        self.assertEqual(code, 2)

    def test_a_broken_plumbing_check_exits_two(self) -> None:
        baseline_path, selection_path = self.select(
            fetch_status={"fetched": False, "reason": "no prior successful full run"}
        )
        comparison_path, comparator_status = self.compare(baseline_path)
        comparison = self._read(comparison_path)
        comparison["summary"]["compared_metrics"] = 19
        self._write(comparison_path, comparison)
        code, classification, _ = self.report(
            selection_path, comparison_path, comparator_status
        )
        self.assertEqual(code, 2)
        self.assertEqual(
            classification["outcome"], perf_smoke_baseline.OUTCOME_PLUMBING_DEFECT
        )


class WorkflowWiringTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.workflow = WORKFLOW_PATH.read_text(encoding="utf-8")

    def test_the_workflow_parses_as_yaml(self) -> None:
        try:
            import yaml
        except ImportError:  # pragma: no cover - PyYAML is optional on the runner
            self.skipTest("PyYAML is not installed")
        document = yaml.safe_load(self.workflow)
        self.assertIn("jobs", document)
        self.assertIn("smoke", document["jobs"])
        self.assertIn("full", document["jobs"])

    def test_both_cargo_jobs_declare_the_runner_labels_they_run_on(self) -> None:
        for job in ("smoke", "full"):
            block = self._job(job)
            runs_on = re.search(r"^    runs-on:\s*(\S+)\s*$", block, re.MULTILINE)
            self.assertIsNotNone(runs_on, job)
            labels = re.search(
                r"^      LITCHI_RUNNER_LABELS:\s*(\S+)\s*$", block, re.MULTILINE
            )
            self.assertIsNotNone(labels, job)
            self.assertEqual(labels.group(1), runs_on.group(1), job)

    def test_both_cargo_jobs_point_at_the_checked_smoke_policy(self) -> None:
        for job in ("smoke", "full"):
            self.assertIn(
                "LITCHI_SMOKE_POLICY: docs/performance/perf-smoke-baseline-policy-v1.json",
                self._job(job),
            )

    def test_the_full_job_publishes_the_reference_report_and_descriptor(self) -> None:
        block = self._job("full")
        self.assertIn("--bin litchi-perf-baseline-alloc", block)
        self.assertIn("target/perf/allocator-baseline.json", block)
        self.assertIn("target/perf/allocator-baseline-descriptor.json", block)
        self.assertIn("tools/perf_smoke_baseline.py descriptor", block)

    def test_the_smoke_job_fetches_selects_compares_and_reports(self) -> None:
        block = self._job("smoke")
        for fragment in (
            "gh run list",
            "gh run download",
            "tools/perf_smoke_baseline.py choose-runs",
            "tools/perf_smoke_baseline.py select",
            "tools/perf_compare.py",
            "tools/perf_smoke_baseline.py report",
        ):
            self.assertIn(fragment, block, fragment)

    def test_the_smoke_job_never_lets_the_fetch_fail_the_run(self) -> None:
        block = self._job("smoke")
        fetch = next(
            step
            for step in block.split("\n      - name: ")
            if step.startswith("Fetch the last successful")
        )
        self.assertIn("continue-on-error: true", fetch)

    def test_the_comparator_step_captures_its_status_instead_of_failing(self) -> None:
        block = self._job("smoke")
        self.assertIn("comparator-exit-status.txt", block)
        self.assertIn("--comparator-exit-status", block)

    def test_the_new_tooling_is_in_both_path_filters(self) -> None:
        for event in ("push", "pull_request"):
            section = self._top_level_section(event)
            for path in (
                "tools/perf_smoke_baseline.py",
                "tools/test_perf_smoke_baseline.py",
                "docs/performance/perf-smoke-baseline-policy-v1.json",
            ):
                self.assertRegex(
                    section,
                    rf"(?m)^\s*-\s*'{re.escape(path)}'\s*$",
                    msg=f"{event} filter must include {path}",
                )

    def test_the_new_suite_runs_in_the_smoke_job(self) -> None:
        self.assertIn("tools.test_perf_smoke_baseline", self._job("smoke"))

    def _job(self, name: str) -> str:
        lines = self.workflow.splitlines()
        start = lines.index(f"  {name}:")
        end = len(lines)
        for index in range(start + 1, len(lines)):
            if re.fullmatch(r"  [A-Za-z0-9_-]+:\s*", lines[index]):
                end = index
                break
        return "\n".join(lines[start:end])

    def _top_level_section(self, name: str) -> str:
        lines = self.workflow.splitlines()
        start = lines.index(f"  {name}:")
        end = len(lines)
        for index in range(start + 1, len(lines)):
            if re.fullmatch(r"  [A-Za-z0-9_-]+:\s*", lines[index]):
                end = index
                break
        return "\n".join(lines[start:end])


class CheckedAllocatorPolicyTests(unittest.TestCase):
    def test_the_checked_policy_pins_the_harness_counter_revision(self) -> None:
        policy = perf_compare.validate_policy(checked_comparator_policy())
        self.assertEqual(
            policy["tool_identity"].get("allocator_counter_revision"),
            "serialized_region_peak_v3",
        )

    def test_the_pinned_revision_matches_the_harness_source(self) -> None:
        source = (
            ROOT / "tools" / "perf-baseline" / "src" / "allocation_metrics.rs"
        ).read_text(encoding="utf-8")
        match = re.search(r'then_some\("([a-z0-9_]+)"\)', source)
        self.assertIsNotNone(match)
        self.assertEqual(
            match.group(1),
            checked_comparator_policy()["tool_identity"]["allocator_counter_revision"],
        )


if __name__ == "__main__":
    unittest.main()
