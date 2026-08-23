#!/usr/bin/env python3
"""Validate the versioned performance-claim registry.

The registry is deliberately small and boring: it records the production
decision separately from the evidence that informed that decision.  The
structural mode is suitable for normal CI and does not require the retained
external evidence packages.  Strict mode additionally opens an evidence root,
recomputes the package/member hashes, and checks the ABBA summary and optional
resource guardrail reports.

Only the Python standard library is used.  Exit status 1 denotes a claim or
policy failure; status 2 denotes malformed input, an unsupported schema, an
unsafe path, or required evidence that could not be read.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import math
import re
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import Any, Iterable


SCHEMA_VERSION = 1
REGISTRY_KIND = "litchi-performance-claim-registry"
SHA256_RE = re.compile(r"[0-9a-f]{64}\Z")
REVISION_RE = re.compile(r"[0-9a-f]{40}\Z")
STATISTICS = ("p50", "mean", "p95", "p99")
ABBA_ROLES = ("a1", "b1", "b2", "a2")
ABBA_ORDER = ("a1_control", "b1_candidate", "b2_candidate", "a2_control")
RESOURCE_LEGS = ("A1", "B1", "B2", "A2")
RESOURCE_PAIRS = ("A1_control_to_B1_candidate", "A2_control_to_B2_candidate")

_MISSING = object()


class ClaimRegistryError(Exception):
    """Base class for a fail-closed registry error."""


class ClaimPolicyError(ClaimRegistryError):
    """A structurally valid registry contains an invalid claim."""


class ClaimInputError(ClaimRegistryError):
    """Registry/evidence input is malformed, unavailable, or unsafe."""


def _reject_constant(value: str) -> Any:
    raise ValueError(f"non-finite JSON constant {value!r}")


def _reject_duplicate_keys(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for key, value in pairs:
        if key in result:
            raise ValueError(f"duplicate JSON object key {key!r}")
        result[key] = value
    return result


def _validate_json_tree(value: Any, location: str) -> None:
    if value is None or isinstance(value, (bool, str, int)):
        return
    if isinstance(value, float):
        if not math.isfinite(value):
            raise ClaimInputError(f"{location} contains a non-finite number")
        return
    if isinstance(value, list):
        for index, item in enumerate(value):
            _validate_json_tree(item, f"{location}[{index}]")
        return
    if isinstance(value, dict):
        for key, item in value.items():
            if not isinstance(key, str):
                raise ClaimInputError(f"{location} has a non-string object key")
            _validate_json_tree(item, f"{location}.{key}")
        return
    raise ClaimInputError(f"{location} contains unsupported JSON value")


def canonical_bytes(value: Any) -> bytes:
    """Return the canonical JSON bytes used by the existing perf tools."""

    _validate_json_tree(value, "JSON")
    try:
        return json.dumps(
            value,
            sort_keys=True,
            separators=(",", ":"),
            allow_nan=False,
        ).encode("utf-8")
    except (TypeError, ValueError, OverflowError) as error:
        raise ClaimInputError(f"cannot canonicalize JSON: {error}") from error


def load_json(path: Path, *, location: str | None = None) -> Any:
    label = location or str(path)
    try:
        raw = path.read_bytes()
    except OSError as error:
        raise ClaimInputError(f"cannot read {label}: {error}") from error
    try:
        text = raw.decode("utf-8")
        value = json.loads(
            text,
            object_pairs_hook=_reject_duplicate_keys,
            parse_constant=_reject_constant,
        )
    except (UnicodeDecodeError, json.JSONDecodeError, ValueError) as error:
        raise ClaimInputError(f"invalid JSON in {label}: {error}") from error
    _validate_json_tree(value, label)
    return value


def sha256_file(path: Path) -> tuple[str, int]:
    digest = hashlib.sha256()
    size = 0
    try:
        with path.open("rb") as source:
            while True:
                chunk = source.read(1024 * 1024)
                if not chunk:
                    break
                size += len(chunk)
                digest.update(chunk)
    except OSError as error:
        raise ClaimInputError(f"cannot hash {path}: {error}") from error
    return digest.hexdigest(), size


def _require_object(value: Any, location: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        raise ClaimInputError(f"{location} must be an object")
    return value


def _require_list(value: Any, location: str) -> list[Any]:
    if not isinstance(value, list):
        raise ClaimInputError(f"{location} must be an array")
    return value


def _require_string(value: Any, location: str, *, nonempty: bool = True) -> str:
    if not isinstance(value, str) or (nonempty and not value):
        raise ClaimInputError(f"{location} must be a non-empty string")
    return value


def _require_bool(value: Any, location: str) -> bool:
    if not isinstance(value, bool):
        raise ClaimInputError(f"{location} must be a boolean")
    return value


def _require_int(value: Any, location: str, *, minimum: int = 0) -> int:
    if isinstance(value, bool) or not isinstance(value, int) or value < minimum:
        raise ClaimInputError(f"{location} must be an integer >= {minimum}")
    return value


def _require_keys(value: dict[str, Any], allowed: set[str], required: set[str], location: str) -> None:
    unknown = sorted(set(value) - allowed)
    if unknown:
        raise ClaimInputError(f"{location} has unknown key(s): {', '.join(unknown)}")
    missing = sorted(required - set(value))
    if missing:
        raise ClaimInputError(f"{location} is missing required key(s): {', '.join(missing)}")


def _require_sha(value: Any, location: str) -> str:
    if not isinstance(value, str) or SHA256_RE.fullmatch(value) is None:
        raise ClaimInputError(f"{location} must be a lowercase SHA-256 digest")
    return value


def _require_revision(value: Any, location: str) -> str:
    if not isinstance(value, str) or REVISION_RE.fullmatch(value) is None:
        raise ClaimInputError(f"{location} must be a 40-character lowercase revision")
    return value


def _validate_relative_path(value: Any, location: str) -> str:
    path = _require_string(value, location)
    candidate = Path(path)
    if candidate.is_absolute() or "\\" in path or "\x00" in path:
        raise ClaimInputError(f"{location} must be a safe relative POSIX path")
    if any(part in {"", ".", ".."} for part in path.split("/")):
        raise ClaimInputError(f"{location} must not contain empty, '.' or '..' path components")
    if path.startswith("/"):
        raise ClaimInputError(f"{location} must be relative")
    return path


def safe_path(root: Path, relative: str, *, location: str, require_exists: bool) -> Path:
    rel = _validate_relative_path(relative, location)
    try:
        root_resolved = root.resolve(strict=True)
    except OSError as error:
        raise ClaimInputError(f"cannot resolve {root}: {error}") from error
    candidate = root / rel
    try:
        resolved = candidate.resolve(strict=require_exists)
    except OSError as error:
        if require_exists:
            raise ClaimInputError(f"{location} does not resolve: {error}") from error
        resolved = candidate.resolve(strict=False)
    try:
        resolved.relative_to(root_resolved)
    except ValueError as error:
        raise ClaimInputError(f"{location} escapes its root") from error
    if require_exists and not resolved.exists():
        raise ClaimInputError(f"{location} does not exist")
    return resolved


def _check_string_list(value: Any, location: str, *, allow_empty: bool = False) -> list[str]:
    values = _require_list(value, location)
    if not values and not allow_empty:
        raise ClaimInputError(f"{location} must not be empty")
    result: list[str] = []
    for index, item in enumerate(values):
        text = _require_string(item, f"{location}[{index}]")
        if "*" in text or "?" in text:
            raise ClaimInputError(f"{location}[{index}] must be an exact value, not a wildcard")
        if text in result:
            raise ClaimInputError(f"{location} contains duplicate value {text!r}")
        result.append(text)
    return result


def _validate_policy(registry: dict[str, Any]) -> dict[str, dict[str, Any]]:
    policies = _require_object(registry.get("policies"), "registry.policies")
    if set(policies) != {"latency-abba-v1", "resource-guardrail-v1"}:
        raise ClaimInputError("registry.policies must contain exactly the v1 latency and resource policies")
    latency = _require_object(policies["latency-abba-v1"], "registry.policies.latency-abba-v1")
    _require_keys(
        latency,
        {"minimum_samples", "order", "drift_ceiling_percent"},
        {"minimum_samples", "order", "drift_ceiling_percent"},
        "registry.policies.latency-abba-v1",
    )
    _require_int(latency["minimum_samples"], "registry.policies.latency-abba-v1.minimum_samples", minimum=1)
    if latency["order"] != list(ABBA_ROLES):
        raise ClaimInputError("latency ABBA order must be [a1, b1, b2, a2]")
    ceilings = _require_object(
        latency["drift_ceiling_percent"],
        "registry.policies.latency-abba-v1.drift_ceiling_percent",
    )
    if set(ceilings) != set(STATISTICS):
        raise ClaimInputError("latency drift ceilings must cover p50, mean, p95 and p99")
    for stat in STATISTICS:
        value = ceilings[stat]
        if isinstance(value, bool) or not isinstance(value, (int, float)) or not math.isfinite(float(value)) or value <= 0:
            raise ClaimInputError(f"invalid latency drift ceiling for {stat}")
    resource = _require_object(policies["resource-guardrail-v1"], "registry.policies.resource-guardrail-v1")
    _require_keys(
        resource,
        {"max_regression_percent", "required_pairings"},
        {"max_regression_percent", "required_pairings"},
        "registry.policies.resource-guardrail-v1",
    )
    max_regression = resource["max_regression_percent"]
    if isinstance(max_regression, bool) or not isinstance(max_regression, (int, float)) or not math.isfinite(float(max_regression)) or max_regression < 0:
        raise ClaimInputError("resource max_regression_percent must be finite and non-negative")
    if resource["required_pairings"] != list(RESOURCE_PAIRS):
        raise ClaimInputError("resource guardrail pairings are not the v1 pairings")
    return {"latency": latency, "resource": resource}


def _validate_scope(scope: Any, location: str) -> dict[str, Any]:
    value = _require_object(scope, location)
    _require_keys(value, {"format", "selectors", "corpora"}, {"format", "selectors", "corpora"}, location)
    _require_string(value["format"], f"{location}.format")
    selectors = _check_string_list(value["selectors"], f"{location}.selectors")
    corpora = _require_list(value["corpora"], f"{location}.corpora")
    if not corpora:
        raise ClaimInputError(f"{location}.corpora must not be empty")
    names: set[str] = set()
    normalized: list[dict[str, Any]] = []
    for index, corpus in enumerate(corpora):
        item_location = f"{location}.corpora[{index}]"
        item = _require_object(corpus, item_location)
        _require_keys(
            item,
            {"name", "archive_sha256", "generator", "shape", "package_format"},
            {"name", "archive_sha256"},
            item_location,
        )
        name = _require_string(item["name"], f"{item_location}.name")
        if name in names:
            raise ClaimInputError(f"{location}.corpora contains duplicate {name!r}")
        names.add(name)
        _require_sha(item["archive_sha256"], f"{item_location}.archive_sha256")
        for key in ("generator", "shape", "package_format"):
            if key in item:
                _require_string(item[key], f"{item_location}.{key}")
        normalized.append(item)
    return {"format": value["format"], "selectors": selectors, "corpora": normalized}


def _validate_evidence(evidence: Any, location: str) -> dict[str, Any]:
    value = _require_object(evidence, location)
    kind = _require_string(value.get("kind"), f"{location}.kind")
    if kind == "abba_package":
        _require_keys(value, {"id", "kind", "relative_path", "manifest", "summary"}, {"id", "kind", "relative_path", "manifest", "summary"}, location)
        manifest = _require_object(value["manifest"], f"{location}.manifest")
        _require_keys(manifest, {"path", "sha256", "schema_version"}, {"path", "sha256", "schema_version"}, f"{location}.manifest")
        summary = _require_object(value["summary"], f"{location}.summary")
        _require_keys(summary, {"path", "sha256", "canonical_sha256", "schema_version"}, {"path", "sha256", "canonical_sha256", "schema_version"}, f"{location}.summary")
        _validate_relative_path(value["relative_path"], f"{location}.relative_path")
        _validate_relative_path(manifest["path"], f"{location}.manifest.path")
        _validate_relative_path(summary["path"], f"{location}.summary.path")
        _require_sha(manifest["sha256"], f"{location}.manifest.sha256")
        _require_sha(summary["sha256"], f"{location}.summary.sha256")
        _require_sha(summary["canonical_sha256"], f"{location}.summary.canonical_sha256")
        if manifest["schema_version"] != 1 or summary["schema_version"] != 1:
            raise ClaimInputError(f"{location} evidence schema must be version 1")
    elif kind == "resource_abba_report":
        _require_keys(value, {"id", "kind", "relative_path", "sha256", "schema_version", "abba_schema_version"}, {"id", "kind", "relative_path", "sha256", "schema_version", "abba_schema_version"}, location)
        _validate_relative_path(value["relative_path"], f"{location}.relative_path")
        _require_sha(value["sha256"], f"{location}.sha256")
        if value["schema_version"] != 1 or value["abba_schema_version"] != 1:
            raise ClaimInputError(f"{location} resource evidence schema must be version 1")
    else:
        raise ClaimInputError(f"{location}.kind is unsupported: {kind!r}")
    _require_string(value["id"], f"{location}.id")
    return value


def _validate_cells(value: Any, location: str) -> list[dict[str, str]]:
    cells = _require_list(value, location)
    normalized: list[dict[str, str]] = []
    seen: set[tuple[str, str, str]] = set()
    for index, cell in enumerate(cells):
        item_location = f"{location}[{index}]"
        item = _require_object(cell, item_location)
        _require_keys(item, {"case", "corpus", "statistic"}, {"case", "corpus", "statistic"}, item_location)
        case = _require_string(item["case"], f"{item_location}.case")
        corpus = _require_string(item["corpus"], f"{item_location}.corpus")
        statistic = _require_string(item["statistic"], f"{item_location}.statistic")
        if statistic not in STATISTICS:
            raise ClaimInputError(f"{item_location}.statistic is unsupported")
        key = (case, corpus, statistic)
        if key in seen:
            raise ClaimInputError(f"{location} contains duplicate cell {key!r}")
        seen.add(key)
        normalized.append({"case": case, "corpus": corpus, "statistic": statistic})
    return normalized


def _validate_claim(claim: Any, location: str, evidence_by_id: dict[str, dict[str, Any]]) -> dict[str, Any]:
    value = _require_object(claim, location)
    _require_keys(
        value,
        {"id", "change_id", "claim_class", "status", "code_state", "reason_codes", "scope", "latency_evidence", "resource_guardrail", "documentation"},
        {"id", "change_id", "claim_class", "status", "code_state", "reason_codes", "scope", "latency_evidence", "documentation"},
        location,
    )
    for key in ("id", "change_id", "claim_class"):
        _require_string(value[key], f"{location}.{key}")
    status = _require_string(value["status"], f"{location}.status")
    code_state = _require_string(value["code_state"], f"{location}.code_state")
    allowed_status = {"rejected", "held", "pending", "landed", "descriptive_only", "correctness_only"}
    if status not in allowed_status:
        raise ClaimInputError(f"{location}.status is unsupported: {status!r}")
    if code_state not in {"not_landed", "landed"}:
        raise ClaimInputError(f"{location}.code_state is unsupported: {code_state!r}")
    if (status == "landed") != (code_state == "landed"):
        raise ClaimPolicyError(f"{location} status/code_state disagree")
    reasons = _check_string_list(value["reason_codes"], f"{location}.reason_codes")
    scope = _validate_scope(value["scope"], f"{location}.scope")
    latency = _require_object(value["latency_evidence"], f"{location}.latency_evidence")
    _require_keys(
        latency,
        {"evidence_id", "allowed_statistics", "accepted_cells", "adverse_both_cells", "accepted_statistics"},
        {"evidence_id", "allowed_statistics", "accepted_cells", "adverse_both_cells"},
        f"{location}.latency_evidence",
    )
    evidence_id = _require_string(latency["evidence_id"], f"{location}.latency_evidence.evidence_id")
    evidence = evidence_by_id.get(evidence_id)
    if evidence is None or evidence.get("kind") != "abba_package":
        raise ClaimInputError(f"{location}.latency_evidence references no ABBA package {evidence_id!r}")
    allowed_stats = _check_string_list(latency["allowed_statistics"], f"{location}.latency_evidence.allowed_statistics")
    if any(stat not in STATISTICS for stat in allowed_stats):
        raise ClaimInputError(f"{location}.latency_evidence.allowed_statistics contains an unsupported statistic")
    accepted_count = _require_int(latency["accepted_cells"], f"{location}.latency_evidence.accepted_cells")
    adverse_count = _require_int(latency["adverse_both_cells"], f"{location}.latency_evidence.adverse_both_cells")
    cells = _validate_cells(latency.get("accepted_statistics", []), f"{location}.latency_evidence.accepted_statistics")
    if len(cells) != accepted_count:
        raise ClaimPolicyError(f"{location}.latency_evidence.accepted_cells does not match accepted_statistics")
    docs = _check_string_list(value["documentation"], f"{location}.documentation")
    for index, path in enumerate(docs):
        _validate_relative_path(path, f"{location}.documentation[{index}]")
    resource = value.get("resource_guardrail")
    if resource is not None:
        resource_value = _require_object(resource, f"{location}.resource_guardrail")
        _require_keys(resource_value, {"required", "evidence_id", "metrics"}, {"required", "metrics"}, f"{location}.resource_guardrail")
        required = _require_bool(resource_value["required"], f"{location}.resource_guardrail.required")
        metrics = _check_string_list(resource_value["metrics"], f"{location}.resource_guardrail.metrics")
        if required and not metrics:
            raise ClaimInputError(f"{location}.resource_guardrail.metrics must not be empty")
        resource_id = resource_value.get("evidence_id")
        if resource_id is not None:
            _require_string(resource_id, f"{location}.resource_guardrail.evidence_id")
            referenced = evidence_by_id.get(resource_id)
            if referenced is not None and referenced.get("kind") != "resource_abba_report":
                raise ClaimInputError(f"{location}.resource_guardrail.evidence_id is not a resource report")
            if referenced is None and not (status in {"held", "pending"} and "resource_guardrail_pending" in reasons):
                raise ClaimInputError(f"{location}.resource_guardrail references missing evidence {resource_id!r}")
        elif required and status == "landed":
            raise ClaimPolicyError(f"{location} landed claim requires a resource evidence_id")
    elif status == "landed":
        # A latency-only landed claim is valid; no resource requirement is implied.
        pass
    if status == "landed" and not cells:
        raise ClaimPolicyError(f"{location} landed claim has no accepted latency cells")
    return {"value": value, "scope": scope, "latency": latency, "accepted_cells": cells, "resource": resource}


def validate_registry(registry: Any, *, repo_root: Path) -> tuple[dict[str, Any], dict[str, dict[str, Any]], dict[str, dict[str, Any]]]:
    value = _require_object(registry, "registry")
    _require_keys(
        value,
        {"schema_version", "registry_kind", "canonicalization", "policies", "evidence", "claims"},
        {"schema_version", "registry_kind", "canonicalization", "policies", "evidence", "claims"},
        "registry",
    )
    if value["schema_version"] != SCHEMA_VERSION or value["registry_kind"] != REGISTRY_KIND:
        raise ClaimInputError("unsupported performance claim registry schema")
    canonical = _require_object(value["canonicalization"], "registry.canonicalization")
    _require_keys(canonical, {"algorithm", "hash"}, {"algorithm", "hash"}, "registry.canonicalization")
    if canonical != {"algorithm": "sorted-json-utf8-compact-v1", "hash": "sha256"}:
        raise ClaimInputError("registry canonicalization is not the v1 algorithm")
    policies = _validate_policy(value)
    evidence_list = _require_list(value["evidence"], "registry.evidence")
    evidence_by_id: dict[str, dict[str, Any]] = {}
    for index, item in enumerate(evidence_list):
        parsed = _validate_evidence(item, f"registry.evidence[{index}]")
        evidence_id = parsed["id"]
        if evidence_id in evidence_by_id:
            raise ClaimInputError(f"duplicate evidence id {evidence_id!r}")
        evidence_by_id[evidence_id] = parsed
    claims_list = _require_list(value["claims"], "registry.claims")
    if not claims_list:
        raise ClaimInputError("registry.claims must not be empty")
    claims_by_id: dict[str, dict[str, Any]] = {}
    for index, item in enumerate(claims_list):
        parsed = _validate_claim(item, f"registry.claims[{index}]", evidence_by_id)
        claim_id = parsed["value"]["id"]
        if claim_id in claims_by_id:
            raise ClaimInputError(f"duplicate claim id {claim_id!r}")
        claims_by_id[claim_id] = parsed
        # Human-readable documentation belongs to the repository, not to an
        # external evidence root.  This is a structural reference check.
        for doc_index, doc in enumerate(parsed["value"]["documentation"]):
            safe_path(repo_root, doc, location=f"registry.claims[{index}].documentation[{doc_index}]", require_exists=True)
    return value, evidence_by_id, claims_by_id


def _decompress_json(path: Path, *, location: str) -> tuple[Any, str, int]:
    """Decompress one report without retaining its raw bytes in Python."""

    try:
        process = subprocess.Popen(
            ["zstd", "-q", "-d", "-c", str(path)],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
        )
    except OSError as error:
        raise ClaimInputError(f"cannot run zstd for {location}: {error}") from error
    digest = hashlib.sha256()
    size = 0
    with process, tempfile.NamedTemporaryFile(prefix="litchi-claim-report-", suffix=".json") as temporary:
        assert process.stdout is not None
        while True:
            chunk = process.stdout.read(1024 * 1024)
            if not chunk:
                break
            digest.update(chunk)
            size += len(chunk)
            temporary.write(chunk)
        stderr = process.stderr.read() if process.stderr is not None else b""
        return_code = process.wait()
        if return_code != 0:
            message = stderr.decode("utf-8", errors="replace").strip()
            raise ClaimInputError(f"zstd failed for {location}: {message}")
        temporary.flush()
        report = load_json(Path(temporary.name), location=location)
    return report, digest.hexdigest(), size


def _validate_report_identity(report: Any, *, role: str, expected_revision: str, minimum_samples: int, location: str) -> None:
    value = _require_object(report, location)
    if value.get("schema_version") != 1:
        raise ClaimInputError(f"{location}.schema_version must be 1")
    tool = _require_object(value.get("tool"), f"{location}.tool")
    if tool.get("name") != "litchi-perf-baseline" or tool.get("version") != "0.1.0":
        raise ClaimInputError(f"{location}.tool is not the release perf harness")
    environment = _require_object(value.get("environment"), f"{location}.environment")
    if environment.get("git_worktree_dirty") is not False:
        raise ClaimInputError(f"{location}.environment.git_worktree_dirty must be false")
    revision = _require_revision(environment.get("git_revision"), f"{location}.environment.git_revision")
    if revision != expected_revision:
        raise ClaimInputError(f"{location}.environment.git_revision does not match {role}")
    configuration = _require_object(value.get("configuration"), f"{location}.configuration")
    samples = _require_int(configuration.get("samples_per_case"), f"{location}.configuration.samples_per_case", minimum=minimum_samples)
    if samples < minimum_samples:
        raise ClaimInputError(f"{location}.configuration.samples_per_case is below policy")
    _require_int(configuration.get("warmup_iterations_per_case"), f"{location}.configuration.warmup_iterations_per_case", minimum=1)
    results = _require_list(value.get("results"), f"{location}.results")
    if not results:
        raise ClaimInputError(f"{location}.results must not be empty")


def _identity_statuses(summary: dict[str, Any], result_count: int) -> None:
    verification = _require_object(summary.get("verification"), "summary.verification")
    for key in ("tool_identity_verified", "configuration_identity_verified", "environment_stable_identity_verified", "case_corpus_identity_verified", "filesystem_evidence_identity_verified", "statistics_recomputed_from_samples"):
        if verification.get(key) is not True:
            raise ClaimInputError(f"summary.verification.{key} must be true")
    if verification.get("result_count") != result_count:
        raise ClaimInputError("summary.verification.result_count does not match results")
    for field in ("source", "sink", "output_sha256", "operation_metrics"):
        counts = _require_object(verification.get(f"{field}_identity"), f"summary.verification.{field}_identity")
        if set(counts) != {"verified_equal", "consistently_absent"}:
            raise ClaimInputError(f"summary.verification.{field}_identity has unexpected counts")
        if counts["verified_equal"] + counts["consistently_absent"] != result_count:
            raise ClaimInputError(f"summary.verification.{field}_identity counts do not add up")
        verified = verification.get(f"{field}_identity_verified", _MISSING)
        # Schema-1 summaries predate the operation-metrics boolean and carry
        # only its equal/absent cardinality map.  The cardinality is the
        # authoritative check for that channel.
        if verified is not _MISSING and verified is not (counts["verified_equal"] == result_count):
            raise ClaimInputError(f"summary.verification.{field}_identity_verified is inconsistent")
    observed_counts = {
        field: {"verified_equal": 0, "consistently_absent": 0}
        for field in ("source", "sink", "output_sha256", "operation_metrics")
    }
    for index, result in enumerate(_require_list(summary.get("results"), "summary.results")):
        item = _require_object(result, f"summary.results[{index}]")
        identity = _require_object(item.get("identity"), f"summary.results[{index}].identity")
        for field in ("source", "sink", "output_sha256", "operation_metrics"):
            status = identity.get(f"{field}_status")
            if status not in {"verified_equal", "consistently_absent"}:
                raise ClaimInputError(f"summary.results[{index}].identity.{field}_status is invalid")
            observed_counts[field][status] += 1
            canonical_key = f"{field}_canonical_json"
            canonical_value = identity.get(canonical_key)
            if status == "consistently_absent" and canonical_value is not None:
                raise ClaimInputError(
                    f"summary.results[{index}].identity.{canonical_key} must be null when absent"
                )
            if status == "verified_equal" and (
                not isinstance(canonical_value, str) or not canonical_value
            ):
                raise ClaimInputError(
                    f"summary.results[{index}].identity.{canonical_key} must be present when equal"
                )
            if field == "operation_metrics":
                continue
            field_value = item.get(field)
            if field_value is None and status == "verified_equal":
                raise ClaimInputError(
                    f"summary.results[{index}].{field} is absent despite verified identity"
                )
            if field_value is not None and status == "consistently_absent":
                raise ClaimInputError(
                    f"summary.results[{index}].{field} is present despite consistently absent identity"
                )
    for field, counts in observed_counts.items():
        expected = _require_object(
            verification.get(f"{field}_identity"),
            f"summary.verification.{field}_identity",
        )
        if counts != expected:
            raise ClaimInputError(
                f"summary.verification.{field}_identity does not match result identities"
            )


def _summary_cells(summary: dict[str, Any]) -> tuple[dict[tuple[str, str, str], dict[str, Any]], set[tuple[str, str, str]]]:
    rows: dict[tuple[str, str, str], dict[str, Any]] = {}
    accepted: set[tuple[str, str, str]] = set()
    adverse: set[tuple[str, str, str]] = set()
    for index, raw in enumerate(_require_list(summary.get("results"), "summary.results")):
        result = _require_object(raw, f"summary.results[{index}]")
        case = _require_string(result.get("case"), f"summary.results[{index}].case")
        corpus = _require_object(result.get("corpus"), f"summary.results[{index}].corpus")
        name = _require_string(corpus.get("name"), f"summary.results[{index}].corpus.name")
        elapsed = _require_object(result.get("elapsed_ns"), f"summary.results[{index}].elapsed_ns")
        for key in ("accepted_statistics", "adverse_both_statistics"):
            stats = _require_list(elapsed.get(key), f"summary.results[{index}].elapsed_ns.{key}")
            for stat_index, stat in enumerate(stats):
                if stat not in STATISTICS:
                    raise ClaimInputError(f"summary.results[{index}].elapsed_ns.{key}[{stat_index}] is invalid")
                cell = (case, name, stat)
                target = accepted if key == "accepted_statistics" else adverse
                if cell in target:
                    raise ClaimInputError(f"summary contains duplicate {key} cell {cell!r}")
                target.add(cell)
        row_key = (case, name, "")
        if row_key in rows:
            raise ClaimInputError(f"summary contains duplicate case/corpus row {case!r}/{name!r}")
        rows[row_key] = {"case": case, "corpus": corpus, "result": result}
    return rows, accepted | {(case, corpus, f"!adverse:{stat}") for case, corpus, stat in adverse}


def _verify_scope_and_cells(
    summary: dict[str, Any],
    claim: dict[str, Any],
    *,
    package_change_id: str,
    policy: dict[str, Any],
) -> None:
    if package_change_id != claim["change_id"]:
        raise ClaimPolicyError(f"claim {claim['id']!r} change_id does not match its package")
    scope = claim["scope"]
    rows, combined = _summary_cells(summary)
    accepted = {cell for cell in combined if not cell[2].startswith("!adverse:")}
    adverse = {(case, corpus, stat[9:]) for case, corpus, stat in combined if stat.startswith("!adverse:")}
    expected_pairs: dict[str, dict[str, Any]] = {item["name"]: item for item in scope["corpora"]}
    scope_pairs: set[tuple[str, str]] = set()
    for key, row in rows.items():
        case, corpus_name, _ = key
        if case not in scope["selectors"] or corpus_name not in expected_pairs:
            raise ClaimPolicyError(f"summary row {case!r}/{corpus_name!r} is outside claim scope")
        scope_pairs.add((case, corpus_name))
        corpus = row["corpus"]
        expected = expected_pairs[corpus_name]
        if corpus.get("archive_sha256") != expected["archive_sha256"]:
            raise ClaimPolicyError(f"corpus hash for {corpus_name!r} does not match registry scope")
        for key_name in ("generator", "shape", "package_format"):
            if key_name in expected and corpus.get(key_name) != expected[key_name]:
                raise ClaimPolicyError(f"corpus {corpus_name!r} {key_name} does not match registry scope")
    for selector in scope["selectors"]:
        if not any(case == selector for case, _ in scope_pairs):
            raise ClaimPolicyError(f"selector {selector!r} has no result row")
    latency = claim["latency"]
    if latency["accepted_cells"] != len(accepted):
        raise ClaimPolicyError(f"claim {claim['id']!r} accepted_cells disagrees with summary")
    if latency["adverse_both_cells"] != len(adverse):
        raise ClaimPolicyError(f"claim {claim['id']!r} adverse_both_cells disagrees with summary")
    declared = {(item["case"], item["corpus"], item["statistic"]) for item in claim["accepted_cells"]}
    if declared != accepted:
        missing = sorted(accepted - declared)
        extra = sorted(declared - accepted)
        raise ClaimPolicyError(f"claim {claim['id']!r} accepted cell set differs (missing={missing!r}, extra={extra!r})")
    allowed = set(latency["allowed_statistics"])
    if not accepted <= {(case, corpus, stat) for case, corpus, stat in accepted if stat in allowed}:
        raise ClaimPolicyError(f"claim {claim['id']!r} lists a statistic outside allowed_statistics")
    if accepted & adverse:
        raise ClaimPolicyError(f"claim {claim['id']!r} claims an adverse-both cell")
    if claim["status"] == "landed" and adverse:
        raise ClaimPolicyError(f"landed claim {claim['id']!r} has adverse-both cells")
    _ = policy


def verify_abba_package(
    evidence: dict[str, Any],
    claim: dict[str, Any],
    *,
    evidence_root: Path,
    policy: dict[str, Any],
) -> dict[str, Any]:
    package_dir = safe_path(evidence_root, evidence["relative_path"], location=f"evidence.{evidence['id']}.relative_path", require_exists=True)
    if not package_dir.is_dir():
        raise ClaimInputError(f"evidence package {package_dir} is not a directory")
    manifest_meta = evidence["manifest"]
    manifest_path = safe_path(package_dir, manifest_meta["path"], location=f"evidence.{evidence['id']}.manifest.path", require_exists=True)
    actual_manifest_sha, _ = sha256_file(manifest_path)
    if actual_manifest_sha != manifest_meta["sha256"]:
        raise ClaimInputError(f"ABBA manifest hash mismatch for {evidence['id']}")
    manifest = load_json(manifest_path, location=str(manifest_path))
    if not isinstance(manifest, dict) or manifest.get("schema_version") != 1 or manifest.get("manifest_kind") != "litchi-perf-abba-artifacts":
        raise ClaimInputError(f"ABBA manifest {evidence['id']} has unsupported schema")
    _require_keys(
        manifest,
        {"artifacts", "change", "change_id", "compression", "manifest_kind", "manifest_path", "schema_version", "self_excluded", "summary", "summary_identity"},
        {"artifacts", "change", "change_id", "compression", "manifest_kind", "manifest_path", "schema_version", "self_excluded", "summary", "summary_identity"},
        f"{evidence['id']}.manifest",
    )
    if manifest.get("manifest_path") != manifest_path.name:
        raise ClaimInputError(f"ABBA manifest {evidence['id']} manifest_path mismatch")
    if manifest.get("change_id") != claim["change_id"] or manifest.get("change") != claim["change_id"]:
        raise ClaimPolicyError(f"ABBA manifest {evidence['id']} change id does not match claim")
    if manifest.get("self_excluded") is not True:
        raise ClaimInputError(f"ABBA manifest {evidence['id']} is not self-excluding")
    summary_meta = evidence["summary"]
    summary_path = safe_path(package_dir, summary_meta["path"], location=f"evidence.{evidence['id']}.summary.path", require_exists=True)
    summary_raw_sha, summary_bytes = sha256_file(summary_path)
    if summary_raw_sha != summary_meta["sha256"]:
        raise ClaimInputError(f"summary hash mismatch for {evidence['id']}")
    summary = load_json(summary_path, location=str(summary_path))
    summary_object = _require_object(summary, f"{evidence['id']}.summary")
    _require_keys(
        summary_object,
        {"environment", "harness_identity", "implementation_identity", "protocol", "report_identity", "results", "schema_version", "tool", "verification"},
        {"environment", "harness_identity", "implementation_identity", "protocol", "report_identity", "results", "schema_version", "tool", "verification"},
        f"{evidence['id']}.summary",
    )
    canonical_sha = hashlib.sha256(canonical_bytes(summary)).hexdigest()
    if canonical_sha != summary_meta["canonical_sha256"]:
        raise ClaimInputError(f"summary canonical hash mismatch for {evidence['id']}")
    manifest_summary = _require_object(manifest.get("summary"), f"{evidence['id']}.manifest.summary")
    _require_keys(
        manifest_summary,
        {"bytes", "canonical_bytes", "canonical_sha256", "path", "report_identity", "result_count", "schema_version", "sha256", "tool"},
        {"bytes", "canonical_bytes", "canonical_sha256", "path", "report_identity", "result_count", "schema_version", "sha256", "tool"},
        f"{evidence['id']}.manifest.summary",
    )
    for key, expected in (("path", summary_meta["path"]), ("sha256", summary_meta["sha256"]), ("canonical_sha256", summary_meta["canonical_sha256"]), ("schema_version", 1)):
        if manifest_summary.get(key) != expected:
            raise ClaimInputError(f"ABBA manifest {evidence['id']} summary.{key} mismatch")
    if manifest_summary.get("bytes") != summary_bytes:
        raise ClaimInputError(f"ABBA manifest {evidence['id']} summary byte count mismatch")
    if manifest_summary.get("canonical_bytes") != len(canonical_bytes(summary)):
        raise ClaimInputError(f"ABBA manifest {evidence['id']} canonical byte count mismatch")
    if manifest.get("summary_identity") != manifest_summary:
        raise ClaimInputError(f"ABBA manifest {evidence['id']} summary identity mismatch")
    if summary.get("schema_version") != 1:
        raise ClaimInputError(f"ABBA summary {evidence['id']} schema must be version 1")
    tool = _require_object(summary.get("tool"), f"{evidence['id']}.summary.tool")
    if tool != {"name": "litchi-perf-abba-summary", "version": "0.1.0"}:
        raise ClaimInputError(f"ABBA summary {evidence['id']} tool identity mismatch")
    implementation = _require_object(summary.get("implementation_identity"), f"{evidence['id']}.summary.implementation_identity")
    control = _require_object(implementation.get("control"), f"{evidence['id']}.summary.implementation_identity.control")
    candidate = _require_object(implementation.get("candidate"), f"{evidence['id']}.summary.implementation_identity.candidate")
    control_revision = _require_revision(control.get("git_revision"), f"{evidence['id']}.control_revision")
    candidate_revision = _require_revision(candidate.get("git_revision"), f"{evidence['id']}.candidate_revision")
    if control_revision == candidate_revision or implementation.get("distinct") is not True:
        raise ClaimInputError(f"ABBA summary {evidence['id']} revisions are not distinct")
    protocol = _require_object(summary.get("protocol"), f"{evidence['id']}.summary.protocol")
    if protocol.get("order") != list(ABBA_ORDER) or protocol.get("statistics") != list(STATISTICS):
        raise ClaimInputError(f"ABBA summary {evidence['id']} protocol order/statistics mismatch")
    if protocol.get("drift_ceiling_percent") != policy["drift_ceiling_percent"]:
        raise ClaimInputError(f"ABBA summary {evidence['id']} drift policy mismatch")
    result_list = _require_list(summary.get("results"), f"{evidence['id']}.summary.results")
    _identity_statuses(summary, len(result_list))
    verification = _require_object(summary["verification"], f"{evidence['id']}.summary.verification")
    report_identity = _require_object(summary.get("report_identity"), f"{evidence['id']}.summary.report_identity")
    if set(report_identity) != set(ABBA_ROLES):
        raise ClaimInputError(f"ABBA summary {evidence['id']} report identity roles mismatch")
    artifacts = _require_list(manifest.get("artifacts"), f"{evidence['id']}.manifest.artifacts")
    if [item.get("role") for item in artifacts if isinstance(item, dict)] != list(ABBA_ROLES) or len(artifacts) != 4:
        raise ClaimInputError(f"ABBA manifest {evidence['id']} must contain exactly a1,b1,b2,a2")
    artifact_by_role: dict[str, dict[str, Any]] = {}
    reports_by_role: dict[str, dict[str, Any]] = {}
    for index, raw in enumerate(artifacts):
        item = _require_object(raw, f"{evidence['id']}.manifest.artifacts[{index}]")
        _require_keys(
            item,
            {"bytes", "canonical_sha256", "compression", "path", "role", "sha256", "uncompressed_bytes", "uncompressed_sha256"},
            {"bytes", "canonical_sha256", "compression", "path", "role", "sha256", "uncompressed_bytes", "uncompressed_sha256"},
            f"{evidence['id']}.manifest.artifacts[{index}]",
        )
        role = item.get("role")
        if role not in ABBA_ROLES or role in artifact_by_role:
            raise ClaimInputError(f"ABBA manifest {evidence['id']} has invalid/duplicate artifact role")
        artifact_by_role[role] = item
        artifact_path = safe_path(package_dir, _require_string(item.get("path"), f"artifact {role}.path"), location=f"artifact {role}.path", require_exists=True)
        compressed_sha, compressed_bytes = sha256_file(artifact_path)
        if compressed_sha != item.get("sha256") or compressed_bytes != item.get("bytes"):
            raise ClaimInputError(f"ABBA artifact {evidence['id']}/{role} compressed identity mismatch")
        expected_revision = control_revision if role in {"a1", "a2"} else candidate_revision
        report, uncompressed_sha, uncompressed_bytes = _decompress_json(artifact_path, location=f"{evidence['id']}/{role}")
        if uncompressed_sha != item.get("uncompressed_sha256") or uncompressed_bytes != item.get("uncompressed_bytes"):
            raise ClaimInputError(f"ABBA artifact {evidence['id']}/{role} uncompressed identity mismatch")
        report_canonical = hashlib.sha256(canonical_bytes(report)).hexdigest()
        if report_canonical != item.get("canonical_sha256"):
            raise ClaimInputError(f"ABBA artifact {evidence['id']}/{role} canonical identity mismatch")
        summary_report_identity = _require_object(report_identity[role], f"summary.report_identity.{role}")
        if summary_report_identity.get("canonical_sha256") != report_canonical:
            raise ClaimInputError(f"ABBA summary report identity mismatch for {evidence['id']}/{role}")
        _validate_report_identity(report, role=role, expected_revision=expected_revision, minimum_samples=policy["minimum_samples"], location=f"{evidence['id']}/{role}")
        # Retain only the small identity projection.  The decompressed report
        # can be hundreds of MiB (the 0248 package); retaining all four full
        # JSON trees would needlessly multiply strict-mode memory use.
        reports_by_role[role] = {
            "environment": report.get("environment"),
            "configuration": report.get("configuration"),
        }
    environment = _require_object(summary.get("environment"), f"{evidence['id']}.summary.environment")
    environment_legs = _require_object(environment.get("legs"), f"{evidence['id']}.summary.environment.legs")
    if set(environment_legs) != set(ABBA_ROLES):
        raise ClaimInputError(f"ABBA summary {evidence['id']} environment leg set mismatch")
    for role in ABBA_ROLES:
        raw_environment = _require_object(reports_by_role[role].get("environment"), f"{evidence['id']}/{role}.environment")
        if environment_legs[role] != raw_environment:
            raise ClaimInputError(f"ABBA summary {evidence['id']} environment identity mismatch for {role}")
    stable_environment = dict(_require_object(reports_by_role["a1"].get("environment"), f"{evidence['id']}/a1.environment"))
    stable_environment.pop("git_revision", None)
    if environment.get("stable") != stable_environment:
        raise ClaimInputError(f"ABBA summary {evidence['id']} stable environment identity mismatch")
    harness_identity = _require_object(summary.get("harness_identity"), f"{evidence['id']}.summary.harness_identity")
    harness_configuration = harness_identity.get("configuration")
    if harness_configuration is not None:
        first_configuration = reports_by_role["a1"].get("configuration")
        if harness_configuration != first_configuration:
            raise ClaimInputError(f"ABBA summary {evidence['id']} harness configuration identity mismatch")
    if verification.get("result_count") != len(result_list):
        raise ClaimInputError(f"ABBA summary {evidence['id']} result count mismatch")
    _verify_scope_and_cells(summary, claim, package_change_id=manifest["change_id"], policy=policy)
    return {
        "summary": summary,
        "control_revision": control_revision,
        "candidate_revision": candidate_revision,
        "result_count": len(result_list),
    }


def _finite_nonnegative(value: Any, location: str) -> float:
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        raise ClaimInputError(f"{location} must be numeric")
    number = float(value)
    if not math.isfinite(number) or number < 0:
        raise ClaimInputError(f"{location} must be finite and non-negative")
    return number


def _finite_number(value: Any, location: str) -> float:
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        raise ClaimInputError(f"{location} must be numeric")
    number = float(value)
    if not math.isfinite(number):
        raise ClaimInputError(f"{location} must be finite")
    return number


def _find_report_value(value: Any, path: str) -> Any:
    pieces = path.split(".")
    if len(pieces) != 2:
        return _MISSING
    root, key = pieces
    if root == "time":
        source = value.get("time") if isinstance(value, dict) else None
        parsed = source.get("parsed") if isinstance(source, dict) else None
    elif root == "heaptrack":
        source = value.get("heaptrack") if isinstance(value, dict) else None
        parsed = source.get("parsed") if isinstance(source, dict) else None
        if parsed is None and isinstance(source, dict):
            printed = source.get("print")
            parsed = printed.get("parsed") if isinstance(printed, dict) else None
    else:
        return _MISSING
    return parsed.get(key, _MISSING) if isinstance(parsed, dict) else _MISSING


def verify_resource_report(
    evidence: dict[str, Any],
    claim: dict[str, Any],
    *,
    evidence_root: Path,
    resource_policy: dict[str, Any],
    latency_result: dict[str, Any] | None,
) -> None:
    report_path = safe_path(evidence_root, evidence["relative_path"], location=f"evidence.{evidence['id']}.relative_path", require_exists=True)
    actual_sha, _ = sha256_file(report_path)
    if actual_sha != evidence["sha256"]:
        raise ClaimInputError(f"resource report hash mismatch for {evidence['id']}")
    report = load_json(report_path, location=str(report_path))
    root = _require_object(report, str(report_path))
    if root.get("schema_version") != 1 or root.get("abba_schema_version") != 1:
        raise ClaimInputError(f"resource report {evidence['id']} schema mismatch")
    legs = _require_list(root.get("legs"), f"{evidence['id']}.legs")
    labels = [item.get("leg") if isinstance(item, dict) else None for item in legs]
    if labels != list(RESOURCE_LEGS):
        raise ClaimInputError(f"resource report {evidence['id']} leg order mismatch")
    validation = _require_object(root.get("validation"), f"{evidence['id']}.validation")
    if validation.get("status") != "validated":
        raise ClaimInputError(f"resource report {evidence['id']} validation did not pass")
    control_revision = _require_revision(validation.get("control_revision"), f"{evidence['id']}.validation.control_revision")
    candidate_revision = _require_revision(validation.get("candidate_revision"), f"{evidence['id']}.validation.candidate_revision")
    if control_revision == candidate_revision:
        raise ClaimInputError(f"resource report {evidence['id']} revisions are identical")
    for index, leg in enumerate(legs):
        item = _require_object(leg, f"{evidence['id']}.legs[{index}]")
        harness = item.get("harness_report") or item.get("report")
        if isinstance(harness, dict):
            environment = harness.get("environment")
            if isinstance(environment, dict) and environment.get("git_worktree_dirty") is not False:
                raise ClaimInputError(f"resource report {evidence['id']} leg {labels[index]} is dirty")
        binary = _require_object(item.get("binary_identity"), f"{evidence['id']}.legs[{index}].binary_identity")
        _require_sha(binary.get("binary_sha256"), f"{evidence['id']}.legs[{index}].binary_identity.binary_sha256")
    binaries = [item["binary_identity"]["binary_sha256"] for item in legs]
    if binaries[0] != binaries[3] or binaries[1] != binaries[2] or binaries[0] == binaries[1]:
        raise ClaimInputError(f"resource report {evidence['id']} binary identities are inconsistent")
    statistics = _require_object(root.get("statistics"), f"{evidence['id']}.statistics")
    if statistics.get("status") != "observed":
        raise ClaimInputError(f"resource report {evidence['id']} statistics are not observed")
    metrics = _require_object(statistics.get("metrics"), f"{evidence['id']}.statistics.metrics")
    guardrail = claim.get("resource") or {}
    required_metrics = guardrail.get("metrics", [])
    max_regression = float(resource_policy["max_regression_percent"])
    for metric in required_metrics:
        item = metrics.get(metric)
        if not isinstance(item, dict):
            raise ClaimInputError(f"resource report {evidence['id']} is missing metric {metric!r}")
        paired = _require_object(item.get("paired"), f"{evidence['id']}.statistics.metrics.{metric}.paired")
        for pairing in RESOURCE_PAIRS:
            pair = _require_object(paired.get(pairing), f"{evidence['id']}.statistics.metrics.{metric}.paired.{pairing}")
            if pair.get("status") != "observed":
                raise ClaimPolicyError(f"resource metric {metric} {pairing} is withheld")
            control = _finite_nonnegative(pair.get("control"), f"{metric}.{pairing}.control")
            candidate = _finite_nonnegative(pair.get("candidate"), f"{metric}.{pairing}.candidate")
            if control == 0:
                raise ClaimPolicyError(f"resource metric {metric} {pairing} has zero control")
            relative = _finite_number(pair.get("relative_delta_percent"), f"{metric}.{pairing}.relative_delta_percent")
            if relative > max_regression:
                raise ClaimPolicyError(f"resource metric {metric} {pairing} exceeds +{max_regression:g}%")
            if candidate < 0 or control < 0:
                raise ClaimInputError(f"resource metric {metric} has negative value")
    if latency_result is not None:
        if validation.get("control_revision") != latency_result["control_revision"] or validation.get("candidate_revision") != latency_result["candidate_revision"]:
            raise ClaimPolicyError(f"resource report {evidence['id']} revisions do not match latency evidence")
        expected_corpus = latency_result["summary"].get("results", [{}])[0].get("corpus", {}).get("archive_sha256")
        serialized = json.dumps(root, sort_keys=True)
        if expected_corpus and expected_corpus not in serialized:
            raise ClaimPolicyError(f"resource report {evidence['id']} does not carry the latency corpus identity")


def lint_registry(
    registry_path: Path,
    *,
    repo_root: Path,
    evidence_root: Path | None,
    mode: str,
) -> tuple[int, list[str]]:
    try:
        registry = load_json(registry_path, location=str(registry_path))
        _, evidence_by_id, claims_by_id = validate_registry(registry, repo_root=repo_root)
        for claim_id, parsed in claims_by_id.items():
            claim = parsed["value"]
            if claim["status"] == "landed" and mode == "structural":
                raise ClaimInputError(
                    f"landed claim {claim_id!r} requires strict evidence verification"
                )
            for evidence_id in (parsed["latency"]["evidence_id"],):
                evidence = evidence_by_id[evidence_id]
                if mode == "strict" and evidence_root is not None:
                    verify_abba_package(
                        evidence,
                        {**claim, "scope": parsed["scope"], "latency": parsed["latency"], "accepted_cells": parsed["accepted_cells"], "resource": parsed["resource"]},
                        evidence_root=evidence_root,
                        policy=registry["policies"]["latency-abba-v1"],
                    )
            resource = parsed["resource"]
            if resource is not None and resource.get("required") and resource.get("evidence_id") in evidence_by_id:
                if mode == "strict" and evidence_root is not None:
                    latency_result = verify_abba_package(
                        evidence_by_id[parsed["latency"]["evidence_id"]],
                        {**claim, "scope": parsed["scope"], "latency": parsed["latency"], "accepted_cells": parsed["accepted_cells"], "resource": parsed["resource"]},
                        evidence_root=evidence_root,
                        policy=registry["policies"]["latency-abba-v1"],
                    )
                    verify_resource_report(
                        evidence_by_id[resource["evidence_id"]],
                        {**claim, "resource": resource},
                        evidence_root=evidence_root,
                        resource_policy=registry["policies"]["resource-guardrail-v1"],
                        latency_result=latency_result,
                    )
            elif claim["status"] == "landed" and mode == "strict":
                raise ClaimInputError(f"landed claim {claim_id!r} is missing required resource evidence")
        return 0, [f"OK: {len(claims_by_id)} performance claims validated ({mode})"]
    except ClaimPolicyError as error:
        return 1, [f"CLAIM POLICY ERROR: {error}"]
    except (ClaimInputError, OSError, ValueError, TypeError) as error:
        return 2, [f"INVALID CLAIM REGISTRY: {error}"]


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--registry", type=Path, required=True)
    parser.add_argument("--repo-root", type=Path, required=True)
    parser.add_argument("--evidence-root", type=Path)
    parser.add_argument("--mode", choices=("structural", "strict"), default="structural")
    parser.add_argument("--canonical-sha256", action="store_true", help="print the canonical registry SHA-256")
    return parser


def main(argv: Iterable[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    repo_root = args.repo_root.resolve()
    evidence_root = args.evidence_root.resolve() if args.evidence_root is not None else None
    if args.mode == "strict" and evidence_root is None:
        print("INVALID CLAIM REGISTRY: strict mode requires --evidence-root", file=sys.stderr)
        return 2
    status, messages = lint_registry(args.registry, repo_root=repo_root, evidence_root=evidence_root, mode=args.mode)
    for message in messages:
        print(message, file=sys.stderr if status else sys.stdout)
    if args.canonical_sha256 and status == 0:
        try:
            registry = load_json(args.registry)
            print(hashlib.sha256(canonical_bytes(registry)).hexdigest())
        except ClaimRegistryError as error:
            print(f"INVALID CLAIM REGISTRY: {error}", file=sys.stderr)
            return 2
    return status


if __name__ == "__main__":
    raise SystemExit(main())
