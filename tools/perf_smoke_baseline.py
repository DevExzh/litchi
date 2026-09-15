"""Choose, label and report the baseline for the CI performance smoke check.

The `Performance baseline` workflow's `smoke` job runs on every qualifying push
and pull request. Until change 0626 it built its comparison baseline as a
byte-copy of the report it had just produced, changed only the recorded git
revision, and asserted that `tools/perf_compare.py` reported `pass`. That is a
working plumbing check for the comparator, but it cannot detect a regression:
the two sides are the same measurement.

This module wires the smoke job to real history. The `full` job already
uploads its release reports as a run artifact; it now also produces the same
bounded allocator report the smoke job produces, plus a descriptor naming the
run. The smoke job lists recent successful `full` runs, downloads the newest
artifact that carries such a report, and compares against it when — and only
when — the fetched report is bound to the same comparator policy identity, the
same case/corpus key manifest digest, the same harness binary profile and the
same runner labels. When no compatible artifact exists the job falls back to
the old self-comparison, which is retained here and is explicitly labelled as a
plumbing check rather than a regression detector.

Every branch of that decision lives in this module rather than in a workflow
heredoc, so `tools/test_perf_smoke_baseline.py` can exercise it.

Enforcement is deliberately advisory. `docs/GOAL.md` deliverable 8 forbids
making noisy cloud-hosted microbenchmarks a hard merge gate until variance is
understood, so a fetched-reference comparison annotates the run and never fails
it unless the smoke policy document says `enforcement: blocking`. The one
outcome that always fails is a broken plumbing check: if the comparator cannot
compare a report with itself, the tooling — not the measurement — is wrong.

Subcommands:

``descriptor``
    Write the reference descriptor beside a report the `full` job uploads.
``fetch-status``
    Record what the reference fetch achieved, so the selector reads a typed
    document instead of interpreting an exit code or a missing directory.
``choose-runs``
    Turn a ``gh run list --json ...`` listing into an ordered list of candidate
    run identifiers. Pure: the caller performs the network access.
``select``
    Decide between the fetched reference and the self-comparison fallback,
    write the chosen baseline document and the machine-readable selection.
``report``
    Classify the comparator's verdict for the chosen mode, render the
    annotations and the job summary, and return the policy's exit status.
"""

from __future__ import annotations

import argparse
import copy
import json
import sys
from pathlib import Path
from typing import Any, Iterable

if __package__:
    from . import perf_compare
else:  # pragma: no cover - exercised by the direct CLI entry point
    import perf_compare


TOOL_NAME = "litchi-perf-smoke-baseline"
TOOL_VERSION = "0.1.0"
SMOKE_POLICY_SCHEMA = 1
DESCRIPTOR_SCHEMA = 1
SELECTION_SCHEMA = 1
CLASSIFICATION_SCHEMA = 1
REPORT_SCHEMA = 1

MODE_FETCHED = "fetched_reference"
MODE_SELF = "self_comparison"

SELF_COMPARISON_REVISION_PREFIX = "allocator-smoke-reference:"

ENFORCEMENT_VALUES = ("advisory", "blocking")

#: Comparator error prefixes that describe the runner rather than the code.
#: A hosted runner can change CPU model, kernel, memory size or toolchain
#: between two runs; the comparator fails closed on those, and that verdict is
#: infrastructure drift, never evidence about a change under review. Any error
#: text outside this set is treated as an input defect, which is the
#: conservative direction: an unrecognised message annotates more loudly, never
#: less.
ENVIRONMENT_DRIFT_PREFIXES = (
    "build identity mismatch for ",
    "reference and current git revisions must differ",
)

OUTCOME_PLUMBING_PASS = "plumbing_pass"
OUTCOME_PLUMBING_DEFECT = "plumbing_defect"
OUTCOME_REFERENCE_PASS = "reference_pass"
OUTCOME_REFERENCE_REGRESSION = "reference_regression"
OUTCOME_REFERENCE_ENVIRONMENT_DRIFT = "reference_environment_drift"
OUTCOME_REFERENCE_INPUT_DEFECT = "reference_input_defect"


class SmokeBaselineError(ValueError):
    """Raised when an input this module must be able to trust is unusable."""


def _require_object(value: Any, location: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        raise SmokeBaselineError(f"{location} must be a JSON object")
    return value


def _require_text(value: Any, location: str) -> str:
    if not isinstance(value, str) or not value:
        raise SmokeBaselineError(f"{location} must be a non-empty string")
    return value


def load_json(path: Path) -> Any:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except OSError as error:
        raise SmokeBaselineError(f"cannot read {path}: {error}") from error
    except json.JSONDecodeError as error:
        raise SmokeBaselineError(f"{path} is not valid JSON: {error}") from error


def normalize_runner_labels(value: Any) -> list[str]:
    """Returns runner labels as an order-insensitive, de-duplicated list."""

    if isinstance(value, str):
        items = [item.strip() for item in value.split(",")]
    elif isinstance(value, list):
        items = []
        for item in value:
            if not isinstance(item, str):
                raise SmokeBaselineError("runner labels must be strings")
            items.append(item.strip())
    else:
        raise SmokeBaselineError("runner labels must be a string or a list")
    labels = sorted({item for item in items if item})
    if not labels:
        raise SmokeBaselineError("runner labels must not be empty")
    return labels


def fetch_status_document(
    *, fetched: bool, run_id: str | None, reason: str
) -> dict[str, Any]:
    """Returns the record the fetch step leaves for the selector.

    The shell that talks to the GitHub API writes exactly one of these, so the
    selector never has to interpret an exit code or a missing directory.
    """

    document: dict[str, Any] = {
        "schema_version": 1,
        "fetched": bool(fetched),
        "reason": _require_text(reason, "fetch status reason"),
    }
    if fetched and not run_id:
        raise SmokeBaselineError("a fetched reference must name its run id")
    if run_id:
        document["run_id"] = _require_text(run_id, "fetch status run id")
    return document


def validate_smoke_policy(raw: Any) -> dict[str, Any]:
    """Validates the smoke baseline policy document."""

    policy = _require_object(raw, "smoke policy")
    required = {
        "schema_version",
        "policy_id",
        "comparator_policy",
        "comparator_policy_id",
        "reference_workflow",
        "reference_branch",
        "reference_events",
        "reference_artifact_prefix",
        "reference_report_name",
        "reference_descriptor_name",
        "max_reference_runs",
        "enforcement",
        "enforcement_reason",
        "require_runner_label_match",
        "require_result_key_manifest_match",
        "self_comparison_expectations",
    }
    missing = sorted(required - policy.keys())
    unknown = sorted(policy.keys() - required)
    if missing or unknown:
        raise SmokeBaselineError(
            f"smoke policy keys mismatch: missing={missing}, unknown={unknown}"
        )
    if policy["schema_version"] != SMOKE_POLICY_SCHEMA:
        raise SmokeBaselineError(
            "smoke policy schema_version must be "
            f"{SMOKE_POLICY_SCHEMA}, got {policy['schema_version']!r}"
        )
    for field in (
        "policy_id",
        "comparator_policy",
        "comparator_policy_id",
        "reference_workflow",
        "reference_branch",
        "reference_artifact_prefix",
        "reference_report_name",
        "reference_descriptor_name",
        "enforcement_reason",
    ):
        _require_text(policy[field], f"smoke policy.{field}")
    events = policy["reference_events"]
    if (
        not isinstance(events, list)
        or not events
        or any(not isinstance(event, str) or not event for event in events)
        or len(set(events)) != len(events)
    ):
        raise SmokeBaselineError(
            "smoke policy.reference_events must contain unique non-empty strings"
        )
    limit = policy["max_reference_runs"]
    if isinstance(limit, bool) or not isinstance(limit, int) or limit < 1:
        raise SmokeBaselineError(
            "smoke policy.max_reference_runs must be a positive integer"
        )
    if policy["enforcement"] not in ENFORCEMENT_VALUES:
        raise SmokeBaselineError(
            "smoke policy.enforcement must be "
            f"one of {list(ENFORCEMENT_VALUES)}, got {policy['enforcement']!r}"
        )
    for field in ("require_runner_label_match", "require_result_key_manifest_match"):
        if not isinstance(policy[field], bool):
            raise SmokeBaselineError(f"smoke policy.{field} must be boolean")
    expectations = _require_object(
        policy["self_comparison_expectations"],
        "smoke policy.self_comparison_expectations",
    )
    expected_fields = {
        "status",
        "matched_results",
        "regressions",
        "latency_claims",
        "latency_compared_results",
        "latency_excluded_results",
        "compared_metrics",
    }
    if set(expectations) != expected_fields:
        raise SmokeBaselineError(
            "smoke policy.self_comparison_expectations must contain exactly "
            f"{sorted(expected_fields)}"
        )
    for field in ("status", "latency_claims"):
        _require_text(
            expectations[field], f"smoke policy.self_comparison_expectations.{field}"
        )
    for field in (
        "matched_results",
        "regressions",
        "latency_compared_results",
        "latency_excluded_results",
        "compared_metrics",
    ):
        value = expectations[field]
        if isinstance(value, bool) or not isinstance(value, int) or value < 0:
            raise SmokeBaselineError(
                f"smoke policy.self_comparison_expectations.{field} must be a "
                "non-negative integer"
            )
    return policy


def validate_descriptor(raw: Any) -> dict[str, Any]:
    """Validates a reference descriptor written beside an uploaded report."""

    descriptor = _require_object(raw, "descriptor")
    required = {
        "schema_version",
        "tool",
        "policy_id",
        "comparator_policy_id",
        "run_id",
        "run_attempt",
        "event",
        "runner_labels",
        "git_revision",
        "harness_binary_profile",
        "result_key_manifest_sha256",
        "report_name",
        "environment",
    }
    missing = sorted(required - descriptor.keys())
    unknown = sorted(descriptor.keys() - required)
    if missing or unknown:
        raise SmokeBaselineError(
            f"descriptor keys mismatch: missing={missing}, unknown={unknown}"
        )
    if descriptor["schema_version"] != DESCRIPTOR_SCHEMA:
        raise SmokeBaselineError(
            f"descriptor schema_version must be {DESCRIPTOR_SCHEMA}"
        )
    for field in (
        "policy_id",
        "comparator_policy_id",
        "run_id",
        "event",
        "git_revision",
        "harness_binary_profile",
        "result_key_manifest_sha256",
        "report_name",
    ):
        _require_text(descriptor[field], f"descriptor.{field}")
    digest = descriptor["result_key_manifest_sha256"]
    if len(digest) != 64 or any(character not in "0123456789abcdef" for character in digest):
        raise SmokeBaselineError(
            "descriptor.result_key_manifest_sha256 must be 64 lowercase hex digits"
        )
    descriptor["runner_labels"] = normalize_runner_labels(descriptor["runner_labels"])
    _require_object(descriptor["tool"], "descriptor.tool")
    _require_object(descriptor["environment"], "descriptor.environment")
    attempt = descriptor["run_attempt"]
    if isinstance(attempt, bool) or not isinstance(attempt, int) or attempt < 1:
        raise SmokeBaselineError("descriptor.run_attempt must be a positive integer")
    return descriptor


def build_descriptor(
    report: Any,
    *,
    smoke_policy: dict[str, Any],
    comparator_policy: dict[str, Any],
    runner_labels: Any,
    run_id: str,
    run_attempt: int,
    event: str,
) -> dict[str, Any]:
    """Returns the descriptor that binds an uploaded report to its run."""

    report_object = _require_object(report, "report")
    environment = _require_object(report_object.get("environment"), "report.environment")
    binary_identity = _require_object(
        report_object.get("binary_identity"), "report.binary_identity"
    )
    profile = _require_text(
        binary_identity.get("profile"), "report.binary_identity.profile"
    )
    expected_profile = comparator_policy["tool_identity"]["profile"]
    if profile != expected_profile:
        raise SmokeBaselineError(
            "report binary profile does not match the comparator policy: "
            f"{profile!r} != {expected_profile!r}"
        )
    digest = perf_compare.report_result_key_manifest_sha256(
        report_object,
        comparator_policy["expected_result_count"],
        result_key_fields=tuple(
            comparator_policy.get("result_key_fields", ("case", "corpus"))
        ),
    )
    descriptor = {
        "schema_version": DESCRIPTOR_SCHEMA,
        "tool": {"name": TOOL_NAME, "version": TOOL_VERSION},
        "policy_id": smoke_policy["policy_id"],
        "comparator_policy_id": comparator_policy["policy_id"],
        "run_id": _require_text(run_id, "run_id"),
        "run_attempt": run_attempt,
        "event": _require_text(event, "event"),
        "runner_labels": normalize_runner_labels(runner_labels),
        "git_revision": _require_text(
            environment.get("git_revision"), "report.environment.git_revision"
        ),
        "harness_binary_profile": profile,
        "result_key_manifest_sha256": digest,
        "report_name": smoke_policy["reference_report_name"],
        "environment": {
            field: environment.get(field)
            for field in comparator_policy["build_identity_fields"]
        },
    }
    return validate_descriptor(descriptor)


def choose_reference_runs(
    runs: Any,
    *,
    current_sha: str | None,
    events: Iterable[str],
    limit: int,
) -> list[dict[str, Any]]:
    """Returns candidate reference runs, newest first.

    A candidate is a successful run of a qualifying event whose head commit is
    not the commit under test: the comparator requires distinct revisions, so a
    run of the same commit could never be a usable reference.
    """

    if not isinstance(runs, list):
        raise SmokeBaselineError("run listing must be a JSON array")
    allowed = set(events)
    candidates: list[dict[str, Any]] = []
    for index, raw_run in enumerate(runs):
        if not isinstance(raw_run, dict):
            raise SmokeBaselineError(f"run listing entry {index} must be an object")
        run_id = raw_run.get("databaseId")
        if isinstance(run_id, bool) or not isinstance(run_id, int) or run_id < 1:
            continue
        if raw_run.get("conclusion") != "success":
            continue
        event = raw_run.get("event")
        if not isinstance(event, str) or event not in allowed:
            continue
        head_sha = raw_run.get("headSha")
        if current_sha and isinstance(head_sha, str) and head_sha == current_sha:
            continue
        created_at = raw_run.get("createdAt")
        candidates.append(
            {
                "run_id": str(run_id),
                "event": event,
                "head_sha": head_sha if isinstance(head_sha, str) else None,
                "created_at": created_at if isinstance(created_at, str) else None,
            }
        )
    candidates.sort(key=lambda item: (item["created_at"] or "", item["run_id"]), reverse=True)
    return candidates[:limit]


def self_comparison_baseline(current: Any) -> dict[str, Any]:
    """Returns today's report relabelled as its own reference.

    This is the pre-0626 fallback, kept verbatim in behaviour: a deep copy of
    the current report whose only difference is a prefixed git revision, so the
    comparator's `require_distinct_revisions` check accepts it. It proves the
    comparator runs; it proves nothing about performance.
    """

    current_object = _require_object(current, "current report")
    environment = _require_object(
        current_object.get("environment"), "current report.environment"
    )
    revision = _require_text(
        environment.get("git_revision"), "current report.environment.git_revision"
    )
    baseline = copy.deepcopy(current_object)
    baseline["environment"]["git_revision"] = (
        f"{SELF_COMPARISON_REVISION_PREFIX}{revision}"
    )
    return baseline


def _check(name: str, passed: bool, detail: str) -> dict[str, Any]:
    return {"name": name, "passed": passed, "detail": detail}


def check_reference(
    *,
    current: dict[str, Any],
    candidate: Any,
    descriptor: Any,
    smoke_policy: dict[str, Any],
    comparator_policy: dict[str, Any],
    runner_labels: list[str],
    fetched_run_id: str | None = None,
) -> list[dict[str, Any]]:
    """Returns one record per compatibility check against a fetched reference.

    The checks are the four bindings the brief requires — policy identity,
    case/corpus key manifest digest, harness binary profile and runner labels —
    plus the two the comparator would otherwise reject outright, a distinct
    revision and a clean reference worktree. Every check runs; the caller
    reports all failures rather than only the first.
    """

    checks: list[dict[str, Any]] = []

    if not isinstance(candidate, dict):
        checks.append(
            _check(
                "reference_report_parsed",
                False,
                "the fetched artifact does not contain a JSON object report",
            )
        )
        return checks
    checks.append(
        _check("reference_report_parsed", True, "the fetched report parsed as an object")
    )

    expected_schema = perf_compare.SUPPORTED_REPORT_SCHEMA
    actual_schema = candidate.get("schema_version")
    checks.append(
        _check(
            "reference_report_schema",
            actual_schema == expected_schema,
            f"reference schema_version {actual_schema!r}, expected {expected_schema!r}",
        )
    )

    tool_identity = comparator_policy["tool_identity"]
    checks.append(
        _check(
            "reference_tool_identity",
            candidate.get("tool") == tool_identity,
            "reference tool identity "
            + ("matches" if candidate.get("tool") == tool_identity else "differs from")
            + f" comparator policy {comparator_policy['policy_id']!r}",
        )
    )
    checks.append(
        _check(
            "current_tool_identity",
            current.get("tool") == tool_identity,
            "current tool identity "
            + ("matches" if current.get("tool") == tool_identity else "differs from")
            + f" comparator policy {comparator_policy['policy_id']!r}",
        )
    )

    expected_profile = tool_identity["profile"]
    candidate_binary = candidate.get("binary_identity")
    candidate_profile = (
        candidate_binary.get("profile") if isinstance(candidate_binary, dict) else None
    )
    current_binary = current.get("binary_identity")
    current_profile = (
        current_binary.get("profile") if isinstance(current_binary, dict) else None
    )
    checks.append(
        _check(
            "harness_binary_profile",
            candidate_profile == expected_profile and current_profile == expected_profile,
            f"reference profile {candidate_profile!r}, current profile "
            f"{current_profile!r}, expected {expected_profile!r}",
        )
    )

    descriptor_error: str | None = None
    parsed_descriptor: dict[str, Any] | None = None
    try:
        parsed_descriptor = validate_descriptor(descriptor)
    except SmokeBaselineError as error:
        descriptor_error = str(error)
    checks.append(
        _check(
            "reference_descriptor",
            parsed_descriptor is not None,
            "descriptor accepted"
            if parsed_descriptor is not None
            else f"descriptor rejected: {descriptor_error}",
        )
    )

    if parsed_descriptor is not None:
        checks.append(
            _check(
                "descriptor_policy_identity",
                parsed_descriptor["policy_id"] == smoke_policy["policy_id"]
                and parsed_descriptor["comparator_policy_id"]
                == comparator_policy["policy_id"],
                "descriptor names "
                f"{parsed_descriptor['policy_id']!r}/"
                f"{parsed_descriptor['comparator_policy_id']!r}; expected "
                f"{smoke_policy['policy_id']!r}/{comparator_policy['policy_id']!r}",
            )
        )
        if smoke_policy["require_runner_label_match"]:
            checks.append(
                _check(
                    "runner_labels",
                    parsed_descriptor["runner_labels"] == runner_labels,
                    f"reference ran on {parsed_descriptor['runner_labels']}, "
                    f"this job runs on {runner_labels}",
                )
            )
        if fetched_run_id is not None:
            checks.append(
                _check(
                    "descriptor_run_identity",
                    parsed_descriptor["run_id"] == fetched_run_id,
                    f"descriptor names run {parsed_descriptor['run_id']!r}; the "
                    f"artifact was downloaded from run {fetched_run_id!r}",
                )
            )
        checks.append(
            _check(
                "descriptor_revision_identity",
                isinstance(candidate.get("environment"), dict)
                and parsed_descriptor["git_revision"]
                == candidate["environment"].get("git_revision"),
                f"descriptor names revision {parsed_descriptor['git_revision']!r}",
            )
        )

    if smoke_policy["require_result_key_manifest_match"]:
        expected_digest = comparator_policy["expected_result_keys_sha256"]
        try:
            actual_digest = perf_compare.report_result_key_manifest_sha256(
                candidate,
                comparator_policy["expected_result_count"],
                result_key_fields=tuple(
                    comparator_policy.get("result_key_fields", ("case", "corpus"))
                ),
            )
        except (perf_compare.ComparisonInputError, ValueError, TypeError) as error:
            actual_digest = None
            digest_detail = f"reference key manifest could not be computed: {error}"
        else:
            digest_detail = (
                f"reference key manifest {actual_digest}, policy {expected_digest}"
            )
        descriptor_digest_agrees = (
            parsed_descriptor is None
            or parsed_descriptor["result_key_manifest_sha256"] == actual_digest
        )
        checks.append(
            _check(
                "corpus_key_manifest_sha256",
                actual_digest is not None
                and actual_digest == expected_digest
                and descriptor_digest_agrees,
                digest_detail
                if descriptor_digest_agrees
                else digest_detail + "; descriptor digest disagrees with the report",
            )
        )

    candidate_environment = candidate.get("environment")
    candidate_revision = (
        candidate_environment.get("git_revision")
        if isinstance(candidate_environment, dict)
        else None
    )
    current_environment = current.get("environment")
    current_revision = (
        current_environment.get("git_revision")
        if isinstance(current_environment, dict)
        else None
    )
    checks.append(
        _check(
            "distinct_revisions",
            isinstance(candidate_revision, str)
            and bool(candidate_revision)
            and candidate_revision != current_revision,
            f"reference revision {candidate_revision!r}, current revision "
            f"{current_revision!r}",
        )
    )
    dirty = (
        candidate_environment.get("git_worktree_dirty")
        if isinstance(candidate_environment, dict)
        else None
    )
    checks.append(
        _check(
            "reference_clean_worktree",
            dirty is False,
            f"reference git_worktree_dirty is {dirty!r}",
        )
    )
    return checks


def select_baseline(
    *,
    current: Any,
    candidate: Any,
    descriptor: Any,
    fetch_status: Any,
    smoke_policy: dict[str, Any],
    comparator_policy: dict[str, Any],
    runner_labels: Any,
    current_run_id: str | None = None,
) -> tuple[dict[str, Any], dict[str, Any]]:
    """Returns the baseline document and the machine-readable selection."""

    current_object = _require_object(current, "current report")
    labels = normalize_runner_labels(runner_labels)
    status = _require_object(fetch_status, "fetch status") if fetch_status is not None else {}
    fetched = bool(status.get("fetched"))
    checks: list[dict[str, Any]] = []
    if not fetched:
        reason = status.get("reason")
        checks.append(
            _check(
                "reference_fetched",
                False,
                _require_text(reason, "fetch status.reason")
                if isinstance(reason, str) and reason
                else "no reference artifact was fetched",
            )
        )
    else:
        checks.append(
            _check(
                "reference_fetched",
                True,
                f"downloaded the artifact of run {status.get('run_id')!r}",
            )
        )
        checks.extend(
            check_reference(
                current=current_object,
                candidate=candidate,
                descriptor=descriptor,
                smoke_policy=smoke_policy,
                comparator_policy=comparator_policy,
                runner_labels=labels,
                fetched_run_id=(
                    str(status["run_id"]) if status.get("run_id") is not None else None
                ),
            )
        )

    failures = [check for check in checks if not check["passed"]]
    use_reference = not failures

    current_environment = current_object.get("environment")
    current_revision = (
        current_environment.get("git_revision")
        if isinstance(current_environment, dict)
        else None
    )

    if use_reference:
        baseline = copy.deepcopy(candidate)
        reference_descriptor = validate_descriptor(descriptor)
        reference = {
            "run_id": reference_descriptor["run_id"],
            "event": reference_descriptor["event"],
            "git_revision": reference_descriptor["git_revision"],
            "runner_labels": reference_descriptor["runner_labels"],
        }
        label = (
            "regression comparison against the last successful full run "
            f"({reference['run_id']}, {reference['event']})"
        )
        mode = MODE_FETCHED
    else:
        baseline = self_comparison_baseline(current_object)
        reference = None
        label = (
            "plumbing check only: today's report compared against itself, "
            "which cannot detect a regression"
        )
        mode = MODE_SELF

    selection = {
        "schema_version": SELECTION_SCHEMA,
        "tool": {"name": TOOL_NAME, "version": TOOL_VERSION},
        "policy_id": smoke_policy["policy_id"],
        "comparator_policy_id": comparator_policy["policy_id"],
        "enforcement": smoke_policy["enforcement"],
        "mode": mode,
        "label": label,
        "is_regression_detector": mode == MODE_FETCHED,
        "reference": reference,
        "current": {
            "run_id": current_run_id or None,
            "git_revision": current_revision,
            "runner_labels": labels,
        },
        "checks": checks,
        "fallback_reasons": [
            f"{check['name']}: {check['detail']}" for check in failures
        ],
    }
    return baseline, selection


def validate_selection(raw: Any) -> dict[str, Any]:
    selection = _require_object(raw, "selection")
    if selection.get("schema_version") != SELECTION_SCHEMA:
        raise SmokeBaselineError(f"selection schema_version must be {SELECTION_SCHEMA}")
    if selection.get("mode") not in (MODE_FETCHED, MODE_SELF):
        raise SmokeBaselineError(f"selection mode {selection.get('mode')!r} is unknown")
    return selection


def _self_comparison_defects(
    comparison: dict[str, Any], expectations: dict[str, Any]
) -> list[str]:
    defects: list[str] = []
    status = comparison.get("status")
    if status != expectations["status"]:
        defects.append(
            f"status {status!r}, expected {expectations['status']!r}"
        )
    for error in comparison.get("errors") or []:
        defects.append(f"comparator error: {error}")
    summary = comparison.get("summary")
    if not isinstance(summary, dict):
        defects.append("comparison has no summary object")
        return defects
    for field in (
        "matched_results",
        "regressions",
        "latency_claims",
        "latency_compared_results",
        "latency_excluded_results",
        "compared_metrics",
    ):
        actual = summary.get(field)
        expected = expectations[field]
        if actual != expected:
            defects.append(f"summary.{field} is {actual!r}, expected {expected!r}")
    return defects


def _is_environment_drift(comparison: dict[str, Any]) -> bool:
    errors = comparison.get("errors")
    if not isinstance(errors, list) or not errors:
        return False
    return all(
        isinstance(error, str)
        and any(error.startswith(prefix) for prefix in ENVIRONMENT_DRIFT_PREFIXES)
        for error in errors
    )


def classify_comparison(
    *,
    comparison: Any,
    comparator_exit_status: int,
    selection: dict[str, Any],
    smoke_policy: dict[str, Any],
) -> dict[str, Any]:
    """Classifies the comparator verdict for the selected comparison mode."""

    comparison_object = _require_object(comparison, "comparison")
    status = comparison_object.get("status")
    if status not in ("pass", "regression", "invalid"):
        raise SmokeBaselineError(f"comparison status {status!r} is unknown")
    mode = selection["mode"]
    detail: list[str] = []

    if mode == MODE_SELF:
        defects = _self_comparison_defects(
            comparison_object, smoke_policy["self_comparison_expectations"]
        )
        detail.extend(defects)
        if comparator_exit_status != 0:
            detail.append(f"comparator exit status {comparator_exit_status}")
        outcome = (
            OUTCOME_PLUMBING_PASS
            if not detail
            else OUTCOME_PLUMBING_DEFECT
        )
    elif status == "pass":
        outcome = OUTCOME_REFERENCE_PASS
        summary = comparison_object.get("summary") or {}
        detail.append(
            f"{summary.get('matched_results')} matched results, "
            f"{summary.get('compared_metrics')} metrics, no regression"
        )
    elif status == "regression":
        outcome = OUTCOME_REFERENCE_REGRESSION
        for item in comparison_object.get("regressions") or []:
            if not isinstance(item, dict):
                continue
            detail.append(
                f"{item.get('case')} {item.get('metric')}: "
                f"{item.get('baseline')} -> {item.get('current')}"
            )
    elif _is_environment_drift(comparison_object):
        outcome = OUTCOME_REFERENCE_ENVIRONMENT_DRIFT
        detail.extend(str(error) for error in comparison_object.get("errors") or [])
    else:
        outcome = OUTCOME_REFERENCE_INPUT_DEFECT
        detail.extend(str(error) for error in comparison_object.get("errors") or [])

    blocking = _blocking(outcome, smoke_policy)
    return {
        "schema_version": CLASSIFICATION_SCHEMA,
        "tool": {"name": TOOL_NAME, "version": TOOL_VERSION},
        "mode": mode,
        "enforcement": smoke_policy["enforcement"],
        "comparator_status": status,
        "comparator_exit_status": comparator_exit_status,
        "outcome": outcome,
        "blocking": blocking,
        "exit_status": exit_status(outcome, smoke_policy),
        "detail": detail,
    }


def _blocking(outcome: str, smoke_policy: dict[str, Any]) -> bool:
    if outcome == OUTCOME_PLUMBING_DEFECT:
        return True
    if smoke_policy["enforcement"] != "blocking":
        return False
    return outcome in (OUTCOME_REFERENCE_REGRESSION, OUTCOME_REFERENCE_INPUT_DEFECT)


def exit_status(outcome: str, smoke_policy: dict[str, Any]) -> int:
    """Returns the process exit status the policy assigns to an outcome.

    A broken plumbing check always fails: the comparator could not compare a
    report with itself, which is a tooling defect and not a measurement. Every
    other non-pass outcome is advisory unless the policy says `blocking`, so a
    hosted-runner comparison annotates the job without gating the merge.
    """

    if outcome == OUTCOME_PLUMBING_DEFECT:
        return 2
    if not _blocking(outcome, smoke_policy):
        return 0
    if outcome == OUTCOME_REFERENCE_REGRESSION:
        return 1
    return 2


def render_annotations(
    selection: dict[str, Any], classification: dict[str, Any]
) -> list[str]:
    """Returns GitHub workflow-command annotations for the outcome."""

    outcome = classification["outcome"]
    level = "notice"
    if outcome in (
        OUTCOME_REFERENCE_REGRESSION,
        OUTCOME_REFERENCE_INPUT_DEFECT,
        OUTCOME_PLUMBING_DEFECT,
    ):
        level = "warning"
    if classification["blocking"]:
        level = "error"
    headline = {
        OUTCOME_PLUMBING_PASS: "Performance smoke: plumbing check only",
        OUTCOME_PLUMBING_DEFECT: "Performance smoke: plumbing check failed",
        OUTCOME_REFERENCE_PASS: "Performance smoke: no regression against the last full run",
        OUTCOME_REFERENCE_REGRESSION: "Performance smoke: regression against the last full run",
        OUTCOME_REFERENCE_ENVIRONMENT_DRIFT: (
            "Performance smoke: reference runner differs, comparison withheld"
        ),
        OUTCOME_REFERENCE_INPUT_DEFECT: "Performance smoke: comparison input rejected",
    }[outcome]
    body = [selection["label"]]
    body.extend(classification["detail"])
    if selection["mode"] == MODE_SELF and selection["fallback_reasons"]:
        body.append("fallback because " + "; ".join(selection["fallback_reasons"]))
    if not classification["blocking"] and outcome != OUTCOME_PLUMBING_PASS:
        body.append(
            "advisory: reported, not enforced "
            "(docs/GOAL.md deliverable 8)"
        )
    text = " | ".join(item for item in body if item)
    escaped = (
        text.replace("%", "%25").replace("\r", "%0D").replace("\n", "%0A")
    )
    return [f"::{level} title={headline}::{escaped}"]


def render_summary(
    selection: dict[str, Any], classification: dict[str, Any]
) -> str:
    """Returns the Markdown job summary for the smoke comparison."""

    lines = ["## Performance smoke comparison", ""]
    lines.append(f"- Mode: `{selection['mode']}` — {selection['label']}")
    lines.append(
        "- Detects regressions: "
        + ("yes" if selection["is_regression_detector"] else "**no**")
    )
    reference = selection.get("reference")
    if reference:
        lines.append(
            f"- Reference: run `{reference.get('run_id')}` at revision "
            f"`{reference.get('git_revision')}`"
        )
    lines.append(f"- Comparator status: `{classification['comparator_status']}`")
    lines.append(f"- Outcome: `{classification['outcome']}`")
    lines.append(
        f"- Enforcement: `{classification['enforcement']}` "
        + ("(this outcome fails the job)" if classification["blocking"] else "(advisory)")
    )
    if classification["detail"]:
        lines.extend(["", "Detail:", ""])
        lines.extend(f"- {item}" for item in classification["detail"])
    if selection["fallback_reasons"]:
        lines.extend(["", "Why the fetched reference was not used:", ""])
        lines.extend(f"- {item}" for item in selection["fallback_reasons"])
    return "\n".join(lines) + "\n"


def _write_json(path: Path, document: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        json.dumps(document, indent=2, sort_keys=True, allow_nan=False) + "\n",
        encoding="utf-8",
    )


def _write_text(path: Path, contents: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(contents, encoding="utf-8")


def _append_text(path: Path, contents: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("a", encoding="utf-8") as handle:
        handle.write(contents)


def _optional_json(path: Path | None) -> Any:
    if path is None or not path.exists():
        return None
    return load_json(path)


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    subparsers = parser.add_subparsers(dest="command", required=True)

    descriptor = subparsers.add_parser(
        "descriptor", help="write the reference descriptor for an uploaded report"
    )
    descriptor.add_argument("--policy", type=Path, required=True)
    descriptor.add_argument("--comparator-policy", type=Path, required=True)
    descriptor.add_argument("--report", type=Path, required=True)
    descriptor.add_argument("--runner-labels", required=True)
    descriptor.add_argument("--run-id", required=True)
    descriptor.add_argument("--run-attempt", type=int, default=1)
    descriptor.add_argument("--event", required=True)
    descriptor.add_argument("--out", type=Path, required=True)

    fetch_status = subparsers.add_parser(
        "fetch-status", help="record what the reference fetch achieved"
    )
    fetch_status.add_argument("--out", type=Path, required=True)
    fetch_status.add_argument("--reason", required=True)
    fetch_status.add_argument("--run-id", default="")
    fetch_status.add_argument("--fetched", action="store_true")

    choose = subparsers.add_parser(
        "choose-runs", help="rank candidate reference runs from a gh run listing"
    )
    choose.add_argument("--policy", type=Path, required=True)
    choose.add_argument("--runs", type=Path, required=True)
    choose.add_argument("--current-sha", default="")
    choose.add_argument("--out", type=Path, required=True)

    select = subparsers.add_parser(
        "select", help="choose the fetched reference or the self-comparison fallback"
    )
    select.add_argument("--policy", type=Path, required=True)
    select.add_argument("--comparator-policy", type=Path, required=True)
    select.add_argument("--current", type=Path, required=True)
    select.add_argument("--reference-dir", type=Path)
    select.add_argument("--fetch-status", type=Path)
    select.add_argument("--runner-labels", required=True)
    select.add_argument("--current-run-id", default="")
    select.add_argument("--baseline-out", type=Path, required=True)
    select.add_argument("--selection-out", type=Path, required=True)

    report = subparsers.add_parser(
        "report", help="classify and annotate the comparator verdict"
    )
    report.add_argument("--policy", type=Path, required=True)
    report.add_argument("--selection", type=Path, required=True)
    report.add_argument("--comparison", type=Path, required=True)
    report.add_argument("--comparator-exit-status", type=int, required=True)
    report.add_argument("--classification-out", type=Path)
    report.add_argument("--summary-out", type=Path)
    report.add_argument("--step-summary", type=Path)
    return parser


def _command_descriptor(args: argparse.Namespace) -> int:
    smoke_policy = validate_smoke_policy(load_json(args.policy))
    comparator_policy = perf_compare.validate_policy(load_json(args.comparator_policy))
    descriptor = build_descriptor(
        load_json(args.report),
        smoke_policy=smoke_policy,
        comparator_policy=comparator_policy,
        runner_labels=args.runner_labels,
        run_id=args.run_id,
        run_attempt=args.run_attempt,
        event=args.event,
    )
    _write_json(args.out, descriptor)
    print(
        f"reference descriptor for run {descriptor['run_id']} "
        f"at {descriptor['git_revision']}"
    )
    return 0


def _command_fetch_status(args: argparse.Namespace) -> int:
    document = fetch_status_document(
        fetched=args.fetched, run_id=args.run_id or None, reason=args.reason
    )
    _write_json(args.out, document)
    print(f"reference fetch: {document['reason']}")
    return 0


def _command_choose_runs(args: argparse.Namespace) -> int:
    smoke_policy = validate_smoke_policy(load_json(args.policy))
    candidates = choose_reference_runs(
        load_json(args.runs),
        current_sha=args.current_sha or None,
        events=smoke_policy["reference_events"],
        limit=smoke_policy["max_reference_runs"],
    )
    _write_text(
        args.out, "".join(f"{candidate['run_id']}\n" for candidate in candidates)
    )
    print(f"{len(candidates)} candidate reference run(s)")
    return 0


def _command_select(args: argparse.Namespace) -> int:
    smoke_policy = validate_smoke_policy(load_json(args.policy))
    comparator_policy = perf_compare.validate_policy(load_json(args.comparator_policy))
    current = load_json(args.current)
    reference_dir: Path | None = args.reference_dir
    candidate_path = (
        reference_dir / smoke_policy["reference_report_name"]
        if reference_dir is not None
        else None
    )
    descriptor_path = (
        reference_dir / smoke_policy["reference_descriptor_name"]
        if reference_dir is not None
        else None
    )
    try:
        candidate = _optional_json(candidate_path)
    except SmokeBaselineError as error:
        candidate = {"__unreadable__": str(error)}
    try:
        descriptor = _optional_json(descriptor_path)
    except SmokeBaselineError as error:
        descriptor = {"__unreadable__": str(error)}
    try:
        fetch_status = _optional_json(args.fetch_status)
    except SmokeBaselineError as error:
        fetch_status = {"fetched": False, "reason": f"unreadable fetch status: {error}"}
    if not isinstance(fetch_status, dict):
        fetch_status = {"fetched": False, "reason": "no fetch status was recorded"}
    if candidate is None and fetch_status.get("fetched"):
        fetch_status = {
            "fetched": False,
            "reason": (
                "the fetched artifact carries no "
                f"{smoke_policy['reference_report_name']}"
            ),
        }
    baseline, selection = select_baseline(
        current=current,
        candidate=candidate,
        descriptor=descriptor,
        fetch_status=fetch_status,
        smoke_policy=smoke_policy,
        comparator_policy=comparator_policy,
        runner_labels=args.runner_labels,
        current_run_id=args.current_run_id or None,
    )
    _write_json(args.baseline_out, baseline)
    _write_json(args.selection_out, selection)
    print(f"{selection['mode']}: {selection['label']}")
    for reason in selection["fallback_reasons"]:
        print(f"  fallback reason: {reason}")
    return 0


def _command_report(args: argparse.Namespace) -> int:
    smoke_policy = validate_smoke_policy(load_json(args.policy))
    selection = validate_selection(load_json(args.selection))
    classification = classify_comparison(
        comparison=load_json(args.comparison),
        comparator_exit_status=args.comparator_exit_status,
        selection=selection,
        smoke_policy=smoke_policy,
    )
    summary = render_summary(selection, classification)
    if args.classification_out is not None:
        _write_json(args.classification_out, classification)
    if args.summary_out is not None:
        _write_text(args.summary_out, summary)
    if args.step_summary is not None:
        _append_text(args.step_summary, summary)
    for annotation in render_annotations(selection, classification):
        print(annotation)
    print(summary, end="")
    return classification["exit_status"]


_COMMANDS = {
    "descriptor": _command_descriptor,
    "fetch-status": _command_fetch_status,
    "choose-runs": _command_choose_runs,
    "select": _command_select,
    "report": _command_report,
}


def main(argv: Iterable[str] | None = None) -> int:
    args = _parser().parse_args(list(argv) if argv is not None else None)
    try:
        return _COMMANDS[args.command](args)
    except (
        SmokeBaselineError,
        perf_compare.ComparisonInputError,
        OverflowError,
        TypeError,
        ValueError,
    ) as error:
        print(f"INVALID: {error}", file=sys.stderr)
        return 2
    except OSError as error:
        print(f"INVALID: cannot write output: {error}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
