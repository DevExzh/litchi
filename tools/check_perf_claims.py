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
import statistics
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import Any, Callable, Iterable


SCHEMA_VERSION = 1
REGISTRY_KIND = "litchi-performance-claim-registry"
SHA256_RE = re.compile(r"[0-9a-f]{64}\Z")
REVISION_RE = re.compile(r"[0-9a-f]{40}\Z")
STATISTICS = ("p50", "mean", "p95", "p99")
ABBA_ROLES = ("a1", "b1", "b2", "a2")
ABBA_ORDER = ("a1_control", "b1_candidate", "b2_candidate", "a2_control")
RESOURCE_LEGS = ("A1", "B1", "B2", "A2")
RESOURCE_PAIRS = ("A1_control_to_B1_candidate", "A2_control_to_B2_candidate")
RESOURCE_VARIANTS = {
    "A1": "control",
    "A2": "control",
    "B1": "candidate",
    "B2": "candidate",
}
# These are the retained fields emitted by perf_resource_profile for the
# strict XLSX process-total sources.  Keeping the list here makes the claim
# verifier independent of the publication-side aggregate values.
_TIME_RESOURCE_FIELDS = (
    "max_rss_kib",
    "user_seconds",
    "system_seconds",
    "voluntary_context_switches",
    "involuntary_context_switches",
    "major_page_faults",
    "minor_page_faults",
    "elapsed_wall_seconds",
)
_HEAPTRACK_RESOURCE_FIELDS = (
    "allocation_calls",
    "allocated_bytes",
    "temporary_allocations",
    "peak_heap_bytes",
    "peak_rss_bytes",
)
_MAX_ABBA_DECOMPRESSED_BYTES = 2 * 1024 * 1024 * 1024
_MAX_ABBA_MEMBER_BYTES = 512 * 1024 * 1024
_MAX_ABBA_SUMMARY_BYTES = 64 * 1024 * 1024

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


def canonical_sha256(value: Any) -> str:
    """Hash canonical JSON incrementally to avoid a second large byte buffer."""

    _validate_json_tree(value, "JSON")
    digest = hashlib.sha256()

    def emit(item: Any) -> None:
        if isinstance(item, dict):
            digest.update(b"{")
            for index, key in enumerate(sorted(item)):
                if index:
                    digest.update(b",")
                try:
                    encoded_key = json.dumps(
                        key, separators=(",", ":"), allow_nan=False
                    ).encode("utf-8")
                except (TypeError, ValueError, OverflowError) as error:
                    raise ClaimInputError(f"cannot canonicalize JSON key: {error}") from error
                digest.update(encoded_key)
                digest.update(b":")
                emit(item[key])
            digest.update(b"}")
            return
        if isinstance(item, list):
            digest.update(b"[")
            for index, child in enumerate(item):
                if index:
                    digest.update(b",")
                emit(child)
            digest.update(b"]")
            return
        try:
            encoded = json.dumps(
                item, sort_keys=True, separators=(",", ":"), allow_nan=False
            ).encode("utf-8")
        except (TypeError, ValueError, OverflowError) as error:
            raise ClaimInputError(f"cannot canonicalize JSON value: {error}") from error
        digest.update(encoded)

    emit(value)
    return digest.hexdigest()


def canonical_size(value: Any) -> int:
    """Return canonical JSON size without materializing canonical bytes."""

    _validate_json_tree(value, "JSON")

    def size(item: Any) -> int:
        if isinstance(item, dict):
            total = 2
            for index, key in enumerate(sorted(item)):
                if index:
                    total += 1
                total += len(
                    json.dumps(key, separators=(",", ":"), allow_nan=False).encode(
                        "utf-8"
                    )
                ) + 1 + size(item[key])
            return total
        if isinstance(item, list):
            return 2 + sum(size(child) for child in item) + max(0, len(item) - 1)
        return len(
            json.dumps(item, sort_keys=True, separators=(",", ":"), allow_nan=False).encode(
                "utf-8"
            )
        )

    return size(value)


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


def _decompress_json(
    path: Path,
    *,
    location: str,
    extract: Callable[[Any, str, str], Any] | None = None,
    max_bytes: int | None = None,
) -> tuple[Any, str, int]:
    """Decompress one report without retaining its raw bytes in Python.

    ``extract`` is deliberately called while the temporary JSON tree is still
    local to this function.  Strict verification uses it to retain only the
    source rows and elapsed samples needed for independent recomputation,
    rather than keeping each complete harness report alive alongside the
    summary.
    """

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
            if max_bytes is not None and size > max_bytes:
                try:
                    process.kill()
                except ProcessLookupError:
                    pass
                process.wait()
                raise ClaimInputError(
                    f"{location} exceeds the decompressed-byte ceiling ({max_bytes})"
                )
            temporary.write(chunk)
        stderr = process.stderr.read() if process.stderr is not None else b""
        return_code = process.wait()
        if return_code != 0:
            message = stderr.decode("utf-8", errors="replace").strip()
            raise ClaimInputError(f"zstd failed for {location}: {message}")
        temporary.flush()
        report = load_json(Path(temporary.name), location=location)
    if extract is None:
        extracted = report
    else:
        # The manifest binds the canonical identity of the complete report.
        # Compute it while the decoded tree is local, then discard that tree
        # after the callback returns its bounded projection.
        report_canonical_sha256 = canonical_sha256(report)
        extracted = extract(report, location, report_canonical_sha256)
        if isinstance(extracted, dict):
            extracted["_canonical_sha256"] = report_canonical_sha256
    return extracted, digest.hexdigest(), size


def _canonical_summary_cells(summary: dict[str, Any], *, location: str) -> tuple[set[tuple[str, str, str]], set[tuple[str, str, str]]]:
    """Extract claim cells from a canonical summary."""

    accepted: set[tuple[str, str, str]] = set()
    adverse: set[tuple[str, str, str]] = set()
    for index, result in enumerate(_require_list(summary.get("results"), f"{location}.results")):
        result_location = f"{location}.results[{index}]"
        result_object = _require_object(result, result_location)
        case = _require_string(result_object.get("case"), f"{result_location}.case")
        corpus = _require_object(result_object.get("corpus"), f"{result_location}.corpus")
        corpus_name = _require_string(corpus.get("name"), f"{result_location}.corpus.name")
        elapsed = _require_object(result_object.get("elapsed_ns"), f"{result_location}.elapsed_ns")
        for statistic in _check_string_list(
            elapsed.get("accepted_statistics"),
            f"{result_location}.elapsed_ns.accepted_statistics",
            allow_empty=True,
        ):
            accepted.add((case, corpus_name, statistic))
        for statistic in _check_string_list(
            elapsed.get("adverse_both_statistics"),
            f"{result_location}.elapsed_ns.adverse_both_statistics",
            allow_empty=True,
        ):
            adverse.add((case, corpus_name, statistic))
    return accepted, adverse





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
    recomputed_cells: tuple[set[tuple[str, str, str]], set[tuple[str, str, str]]] | None = None,
) -> None:
    if package_change_id != claim["change_id"]:
        raise ClaimPolicyError(f"claim {claim['id']!r} change_id does not match its package")
    scope = claim["scope"]
    rows, combined = _summary_cells(summary)
    summary_accepted = {cell for cell in combined if not cell[2].startswith("!adverse:")}
    summary_adverse = {
        (case, corpus, stat[9:])
        for case, corpus, stat in combined
        if stat.startswith("!adverse:")
    }
    if recomputed_cells is None:
        accepted = summary_accepted
        adverse = summary_adverse
    else:
        accepted, adverse = recomputed_cells
        if summary_accepted != accepted or summary_adverse != adverse:
            raise ClaimInputError("summary accepted/adverse cells differ from raw samples")
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
    compression = _require_object(manifest.get("compression"), f"{evidence['id']}.manifest.compression")
    if compression.get("format") != "zstd":
        raise ClaimInputError(f"ABBA manifest {evidence['id']} compression format is not zstd")
    if compression.get("checksum") != "XXH64" or compression.get("content_size") is not True:
        raise ClaimInputError(f"ABBA manifest {evidence['id']} compression metadata is incomplete")
    _require_int(compression.get("level"), f"{evidence['id']}.manifest.compression.level", minimum=1)
    _require_int(compression.get("threads"), f"{evidence['id']}.manifest.compression.threads", minimum=1)
    summary_meta = evidence["summary"]
    summary_path = safe_path(package_dir, summary_meta["path"], location=f"evidence.{evidence['id']}.summary.path", require_exists=True)
    summary_raw_sha, summary_bytes = sha256_file(summary_path)
    if summary_raw_sha != summary_meta["sha256"]:
        raise ClaimInputError(f"summary hash mismatch for {evidence['id']}")
    if summary_bytes > _MAX_ABBA_SUMMARY_BYTES:
        raise ClaimInputError(
            f"ABBA summary {evidence['id']} exceeds the byte ceiling ({_MAX_ABBA_SUMMARY_BYTES})"
        )
    summary = load_json(summary_path, location=str(summary_path))
    summary_object = _require_object(summary, f"{evidence['id']}.summary")
    _require_keys(
        summary_object,
        {"environment", "harness_identity", "implementation_identity", "protocol", "report_identity", "results", "schema_version", "tool", "verification"},
        {"environment", "harness_identity", "implementation_identity", "protocol", "report_identity", "results", "schema_version", "tool", "verification"},
        f"{evidence['id']}.summary",
    )
    canonical_sha = canonical_sha256(summary)
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
    if manifest_summary.get("canonical_bytes") != canonical_size(summary):
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
    declared_total = 0
    for index, raw in enumerate(artifacts):
        item = _require_object(raw, f"{evidence['id']}.manifest.artifacts[{index}]")
        declared_bytes = _require_int(
            item.get("uncompressed_bytes"),
            f"artifact {item.get('role', index)}.uncompressed_bytes",
            minimum=0,
        )
        if declared_bytes > _MAX_ABBA_MEMBER_BYTES:
            raise ClaimInputError(
                f"ABBA artifact {evidence['id']}/{item.get('role', index)} exceeds the "
                f"per-member decompressed-byte ceiling ({_MAX_ABBA_MEMBER_BYTES})"
            )
        declared_total += declared_bytes
    if declared_total > _MAX_ABBA_DECOMPRESSED_BYTES:
        raise ClaimInputError(
            f"ABBA evidence {evidence['id']} exceeds the strict decompressed-byte ceiling "
            f"({_MAX_ABBA_DECOMPRESSED_BYTES})"
        )
    try:
        from tools import perf_abba_summary
    except ImportError:  # pragma: no cover - direct script execution fallback
        import perf_abba_summary  # type: ignore[no-redef]
    artifact_by_role: dict[str, dict[str, Any]] = {}
    reports_by_role: dict[str, dict[str, Any]] = {}
    report_profiles: dict[str, str] = {}
    total_uncompressed_bytes = 0
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
        if item.get("compression") != "zstd":
            raise ClaimInputError(f"ABBA artifact {evidence['id']}/{role} compression is not zstd")
        declared_uncompressed_bytes = _require_int(
            item.get("uncompressed_bytes"),
            f"artifact {role}.uncompressed_bytes",
            minimum=0,
        )
        artifact_by_role[role] = item
        artifact_path = safe_path(package_dir, _require_string(item.get("path"), f"artifact {role}.path"), location=f"artifact {role}.path", require_exists=True)
        compressed_sha, compressed_bytes = sha256_file(artifact_path)
        if compressed_sha != item.get("sha256") or compressed_bytes != item.get("bytes"):
            raise ClaimInputError(f"ABBA artifact {evidence['id']}/{role} compressed identity mismatch")
        expected_revision = control_revision if role in {"a1", "a2"} else candidate_revision

        def extract_validated_report(
            value: Any, report_location: str, raw_canonical_sha256: str
        ) -> dict[str, Any]:
            _validate_report_identity(
                value,
                role=role,
                expected_revision=expected_revision,
                minimum_samples=policy["minimum_samples"],
                location=report_location,
            )
            report_profile = perf_abba_summary.detect_report_profile(value, report_location)
            report_profiles[role] = report_profile
            return perf_abba_summary._project_report(
                value,
                report_location,
                profile=report_profile,
                expected_revision=expected_revision,
                minimum_samples=policy["minimum_samples"],
                raw_canonical_sha256=raw_canonical_sha256,
            )

        report, uncompressed_sha, uncompressed_bytes = _decompress_json(
            artifact_path,
            location=f"{evidence['id']}/{role}",
            extract=extract_validated_report,
            max_bytes=_MAX_ABBA_MEMBER_BYTES,
        )
        if uncompressed_sha != item.get("uncompressed_sha256") or uncompressed_bytes != item.get("uncompressed_bytes"):
            raise ClaimInputError(f"ABBA artifact {evidence['id']}/{role} uncompressed identity mismatch")
        total_uncompressed_bytes += uncompressed_bytes
        if total_uncompressed_bytes > _MAX_ABBA_DECOMPRESSED_BYTES:
            raise ClaimInputError(
                f"ABBA evidence {evidence['id']} exceeds the strict decompressed-byte ceiling "
                f"({_MAX_ABBA_DECOMPRESSED_BYTES})"
            )
        report_canonical = report.pop("_canonical_sha256", None)
        if not isinstance(report_canonical, str):
            raise ClaimInputError(f"ABBA artifact {evidence['id']}/{role} lacks canonical identity")
        if report_canonical != item.get("canonical_sha256"):
            raise ClaimInputError(f"ABBA artifact {evidence['id']}/{role} canonical identity mismatch")
        if report.get("report_sha256") != report_canonical:
            raise ClaimInputError(f"ABBA artifact {evidence['id']}/{role} projection identity mismatch")
        summary_report_identity = _require_object(report_identity[role], f"summary.report_identity.{role}")
        if summary_report_identity.get("canonical_sha256") != report_canonical:
            raise ClaimInputError(f"ABBA summary report identity mismatch for {evidence['id']}/{role}")
        reports_by_role[role] = report
    if set(report_profiles) != set(ABBA_ROLES) or len(set(report_profiles.values())) != 1:
        raise ClaimInputError(f"ABBA evidence {evidence['id']} mixes legacy-v1/current-v1 report profiles")
    report_profile = next(iter(report_profiles.values()))

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
    try:
        canonical_summary = perf_abba_summary._summarize_projected_reports(
            reports_by_role,
            drift_ceilings=protocol["drift_ceiling_percent"],
            profile=report_profile,
        )
    except Exception as error:
        raise ClaimInputError(
            f"{evidence['id']} projected reports fail canonical recomputation: {error}"
        ) from error
    if canonical_sha256(canonical_summary) != canonical_sha256(summary):
        raise ClaimInputError(
            f"{evidence['id']} summary differs from canonical raw-report recomputation"
        )
    recomputed_cells = _canonical_summary_cells(
        canonical_summary,
        location=f"{evidence['id']}.canonical",
    )
    _verify_scope_and_cells(
        summary,
        claim,
        package_change_id=manifest["change_id"],
        policy=policy,
        recomputed_cells=recomputed_cells,
    )
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


def _require_resource_artifact(value: Any, location: str) -> dict[str, Any]:
    """Validate a retained artifact descriptor without opening its path.

    Resource reports are evidence packages: the verifier cannot re-read the
    external artifact paths, but it can require the producer's retained
    identity envelope to be complete.  In particular, ``present`` and
    ``retained`` must not be silently replaced by a stale parser value.
    """

    artifact = _require_object(value, location)
    if artifact.get("present") is not True or artifact.get("retained") is not True:
        raise ClaimInputError(f"{location} must be present and retained")
    _require_int(artifact.get("bytes"), f"{location}.bytes")
    _require_sha(artifact.get("sha256"), f"{location}.sha256")
    return artifact


def _require_resource_run(value: Any, location: str) -> dict[str, Any]:
    """Require a completed command and retained stdout/stderr identities."""

    run = _require_object(value, location)
    returncode = run.get("returncode")
    if isinstance(returncode, bool) or not isinstance(returncode, int) or returncode != 0:
        raise ClaimInputError(f"{location} did not complete successfully")
    if run.get("timed_out") is not False:
        raise ClaimInputError(f"{location}.timed_out must be false")
    _require_resource_artifact(run.get("stdout"), f"{location}.stdout")
    _require_resource_artifact(run.get("stderr"), f"{location}.stderr")
    return run


def _validate_time_resource_source(
    source: Any,
    location: str,
    *,
    required_field: str | None = None,
) -> dict[str, Any]:
    """Validate the complete retained /usr/bin/time parser envelope."""

    timed = _require_object(source, location)
    if timed.get("status") != "ok":
        raise ClaimInputError(f"{location}.status must be 'ok'")
    _require_resource_run(timed.get("run"), f"{location}.run")
    parsed = _require_object(timed.get("parsed"), f"{location}.parsed")
    if parsed.get("status") != "ok":
        raise ClaimInputError(f"{location}.parsed.status must be 'ok'")
    _require_resource_artifact(parsed.get("artifact"), f"{location}.parsed.artifact")
    fields = _TIME_RESOURCE_FIELDS
    expected_fields = parsed.get("expected_fields")
    if expected_fields != list(fields):
        raise ClaimInputError(f"{location}.parsed.expected_fields does not match the time schema")
    if required_field is not None and required_field not in fields:
        raise ClaimInputError(f"{location} does not support parsed field {required_field!r}")
    for field in fields:
        _finite_nonnegative(parsed.get(field), f"{location}.parsed.{field}")
    return parsed


def _validate_heaptrack_resource_source(
    source: Any,
    location: str,
    *,
    required_field: str | None = None,
) -> dict[str, Any]:
    """Validate the complete retained heaptrack capture/print envelope."""

    heaptrack = _require_object(source, location)
    if heaptrack.get("status") != "ok":
        raise ClaimInputError(f"{location}.status must be 'ok'")
    _require_resource_artifact(heaptrack.get("harness"), f"{location}.harness")
    _require_resource_run(heaptrack.get("run"), f"{location}.run")
    capture = _require_resource_artifact(heaptrack.get("capture"), f"{location}.capture")
    captures = _require_list(heaptrack.get("captures"), f"{location}.captures")
    if len(captures) != 1:
        raise ClaimInputError(f"{location}.captures must contain exactly one capture")
    listed_capture = _require_resource_artifact(captures[0], f"{location}.captures[0]")
    if listed_capture != capture:
        raise ClaimInputError(f"{location}.captures[0] does not match {location}.capture")

    printed = _require_object(heaptrack.get("print"), f"{location}.print")
    if printed.get("status") != "ok":
        raise ClaimInputError(f"{location}.print.status must be 'ok'")
    print_artifact = _require_resource_artifact(
        printed.get("artifact"), f"{location}.print.artifact"
    )
    print_run = _require_resource_run(printed.get("run"), f"{location}.print.run")
    if print_run.get("stdout") != print_artifact:
        raise ClaimInputError(
            f"{location}.print.run.stdout does not match {location}.print.artifact"
        )
    parsed = _require_object(printed.get("parsed"), f"{location}.print.parsed")
    if parsed.get("status") != "ok":
        raise ClaimInputError(f"{location}.print.parsed.status must be 'ok'")
    parsed_artifact = _require_resource_artifact(
        parsed.get("artifact"), f"{location}.print.parsed.artifact"
    )
    if parsed_artifact != print_artifact:
        raise ClaimInputError(
            f"{location}.print.parsed.artifact does not match {location}.print.artifact"
        )
    _require_resource_artifact(
        parsed.get("histogram_artifact"),
        f"{location}.print.parsed.histogram_artifact",
    )
    fields = _HEAPTRACK_RESOURCE_FIELDS
    if required_field is not None and required_field not in fields:
        raise ClaimInputError(f"{location} does not support parsed field {required_field!r}")
    for field in fields:
        _finite_nonnegative(parsed.get(field), f"{location}.print.parsed.{field}")
    identity = _require_object(
        heaptrack.get("harness_identity"), f"{location}.harness_identity"
    )
    if identity.get("status") != "validated":
        raise ClaimInputError(f"{location}.harness_identity.status must be 'validated'")
    _require_sha(identity.get("sha256"), f"{location}.harness_identity.sha256")
    return parsed


def _parsed_resource_value(leg: dict[str, Any], metric: str) -> Any:
    """Read a resource value only from a successful retained parser source.

    A failed command with a stale ``parsed`` object is an input error, not a
    missing observation.  This distinction is important because callers also
    receive the publication ``values_by_leg`` map, which must be bound to the
    parser result rather than used as an independent fallback.
    """

    if metric.startswith("time."):
        source = leg.get("time")
        if source is None:
            return _MISSING
        field = metric.split(".", 1)[1]
        parsed = _validate_time_resource_source(source, "resource-leg.time", required_field=field)
    elif metric.startswith("heaptrack."):
        source = leg.get("heaptrack")
        if source is None:
            return _MISSING
        field = metric.split(".", 1)[1]
        parsed = _validate_heaptrack_resource_source(
            source, "resource-leg.heaptrack", required_field=field
        )
    else:
        return _MISSING
    return parsed.get(field, _MISSING)


def _compare_resource_number(actual: Any, expected: Any, location: str) -> None:
    actual_value = _finite_number(actual, location)
    expected_value = _finite_number(expected, location)
    tolerance = max(1e-12, abs(expected_value) * 1e-12)
    if abs(actual_value - expected_value) > tolerance:
        raise ClaimInputError(
            f"{location}={actual_value} disagrees with per-leg resource evidence ({expected_value})"
        )


def _resource_value_summary(values: list[float | int], location: str) -> dict[str, Any]:
    observed = [float(value) for value in values]
    try:
        mean = statistics.fmean(observed)
    except (OverflowError, ValueError) as error:
        raise ClaimInputError(f"{location} mean cannot be represented") from error
    try:
        median = statistics.median(observed)
    except (OverflowError, ValueError) as error:
        raise ClaimInputError(f"{location} median cannot be represented") from error
    if not math.isfinite(mean) or not math.isfinite(median):
        raise ClaimInputError(f"{location} aggregate is not finite")
    return {
        "status": "observed",
        "count": len(observed),
        "mean": mean,
        "median": median,
        "minimum": min(observed),
        "maximum": max(observed),
        "overflow_fields": [],
    }


def _resource_pair_summary(
    control: float | int,
    candidate: float | int,
    *,
    execution_order: str,
    location: str,
) -> dict[str, Any]:
    control_value = float(control)
    candidate_value = float(candidate)
    if control_value == 0:
        status = "observed_equal_zero" if candidate_value == 0 else "undefined_zero_control"
        return {
            "execution_order": execution_order,
            "control": control,
            "candidate": candidate,
            "relative_delta_percent": None,
            "ratio_candidate_to_control": None,
            "status": status,
        }
    relative_delta = (candidate_value - control_value) / control_value * 100.0
    ratio = candidate_value / control_value
    if not math.isfinite(relative_delta) or not math.isfinite(ratio):
        raise ClaimInputError(f"{location} pair result is not finite")
    return {
        "execution_order": execution_order,
        "control": control,
        "candidate": candidate,
        "relative_delta_percent": relative_delta,
        "ratio_candidate_to_control": ratio,
        "status": "observed",
    }


def _verify_resource_metric(
    metric_name: str,
    metric: dict[str, Any],
    legs: list[dict[str, Any]],
    *,
    max_regression: float,
    evidence_id: str,
) -> None:
    location = f"{evidence_id}.statistics.metrics.{metric_name}"
    parsed_values: dict[str, float | int] = {}
    for index, leg in enumerate(legs):
        label = RESOURCE_LEGS[index]
        value = _parsed_resource_value(leg, metric_name)
        if value is not _MISSING:
            parsed_values[label] = _finite_nonnegative(
                value, f"{location}.parsed.{label}"
            )

    values_by_leg = metric.get("values_by_leg", _MISSING)
    declared_values: dict[str, float | int] = {}
    if values_by_leg is not _MISSING:
        declared = _require_object(values_by_leg, f"{location}.values_by_leg")
        if set(declared) != set(RESOURCE_LEGS):
            raise ClaimInputError(f"{location}.values_by_leg must contain exactly A1/B1/B2/A2")
        for label in RESOURCE_LEGS:
            declared_values[label] = _finite_nonnegative(
                declared[label], f"{location}.values_by_leg.{label}"
            )

    if set(parsed_values) != set(RESOURCE_LEGS):
        raise ClaimInputError(
            f"{location} requires complete supported parsed source for A1/B1/B2/A2"
        )
    if set(declared_values) != set(RESOURCE_LEGS):
        raise ClaimInputError(
            f"{location}.values_by_leg must contain parsed values for A1/B1/B2/A2"
        )
    for label in RESOURCE_LEGS:
        _compare_resource_number(
            declared_values[label],
            parsed_values[label],
            f"{location}.values_by_leg.{label}",
        )
    values = parsed_values
    controls = [values["A1"], values["A2"]]
    candidates = [values["B1"], values["B2"]]
    expected_control = _resource_value_summary(controls, f"{location}.control")
    expected_candidate = _resource_value_summary(candidates, f"{location}.candidate")
    for group_name, expected in (("control", expected_control), ("candidate", expected_candidate)):
        actual = _require_object(metric.get(group_name), f"{location}.{group_name}")
        for field in ("status", "count", "overflow_fields"):
            if actual.get(field) != expected[field]:
                raise ClaimInputError(f"{location}.{group_name}.{field} disagrees with per-leg evidence")
        for field in ("mean", "median", "minimum", "maximum"):
            _compare_resource_number(actual.get(field), expected[field], f"{location}.{group_name}.{field}")

    paired = _require_object(metric.get("paired"), f"{location}.paired")
    expected_pairs = {
        "A1_control_to_B1_candidate": _resource_pair_summary(
            values["A1"], values["B1"], execution_order="A1, B1", location=f"{location}.paired.A1_control_to_B1_candidate"
        ),
        "A2_control_to_B2_candidate": _resource_pair_summary(
            values["A2"], values["B2"], execution_order="B2, A2", location=f"{location}.paired.A2_control_to_B2_candidate"
        ),
    }
    if set(paired) != set(expected_pairs):
        raise ClaimInputError(f"{location}.paired keys do not match A1/B1/B2/A2")
    for pair_name, expected in expected_pairs.items():
        actual = _require_object(paired.get(pair_name), f"{location}.paired.{pair_name}")
        if actual.get("execution_order") != expected["execution_order"]:
            raise ClaimInputError(f"{location}.paired.{pair_name}.execution_order disagrees")
        for field in ("control", "candidate"):
            _compare_resource_number(actual.get(field), expected[field], f"{location}.paired.{pair_name}.{field}")
        if actual.get("status") != expected["status"]:
            if actual.get("status") != "observed":
                raise ClaimPolicyError(f"resource metric {metric_name} {pair_name} is withheld")
            raise ClaimInputError(f"{location}.paired.{pair_name}.status disagrees with per-leg evidence")
        for field in ("relative_delta_percent", "ratio_candidate_to_control"):
            actual_value = actual.get(field)
            expected_value = expected[field]
            if expected_value is None:
                if actual_value is not None:
                    raise ClaimInputError(f"{location}.paired.{pair_name}.{field} must be null")
            else:
                _compare_resource_number(actual_value, expected_value, f"{location}.paired.{pair_name}.{field}")
        if expected["status"] != "observed":
            raise ClaimPolicyError(f"resource metric {metric_name} {pair_name} is withheld")
        relative = float(expected["relative_delta_percent"])
        if relative > max_regression:
            raise ClaimPolicyError(f"resource metric {metric_name} {pair_name} exceeds +{max_regression:g}%")


def _resource_leg_identity(
    item: dict[str, Any],
    *,
    label: str,
    expected_revision: str,
    expected_variant: str,
    location: str,
) -> tuple[dict[str, Any], dict[str, Any], str]:
    """Validate the identity that binds one published resource leg.

    The resource producer retains the full harness identity beside each leg,
    while a few older reports retained the equivalent ``harness_report``
    environment envelope.  Both forms are accepted, but neither may be
    absent: a leg with only aggregate values is not evidence of a run.

    The returned tuple contains the binary descriptor, the harness tool
    profile, and the retained heaptrack harness-identity digest.  The caller
    compares those values across the ABBA positions below.
    """

    binary = _require_object(item.get("binary_identity"), f"{location}.binary_identity")
    _require_sha(binary.get("binary_sha256"), f"{location}.binary_identity.binary_sha256")
    if binary.get("label") != expected_variant:
        raise ClaimInputError(
            f"{location}.binary_identity.label must be {expected_variant!r}"
        )
    if item.get("variant") != expected_variant:
        raise ClaimInputError(f"{location}.variant must be {expected_variant!r}")
    if item.get("leg") != label:
        raise ClaimInputError(f"{location}.leg must be {label!r}")

    identity = item.get("harness_identity")
    if identity is None:
        # Compatibility for the pre-publication shape.  The environment is
        # still required to carry the revision and clean-worktree assertion.
        identity = item.get("harness_report") or item.get("report")
    identity = _require_object(identity, f"{location}.harness_identity")
    environment = identity.get("environment")
    environment_object = _require_object(environment, f"{location}.harness_identity.environment")
    revision = identity.get("git_revision")
    if revision is None:
        revision = environment_object.get("git_revision")
    revision = _require_revision(revision, f"{location}.harness_identity.git_revision")
    if revision != expected_revision:
        raise ClaimInputError(
            f"{location}.harness_identity.git_revision does not match {expected_variant} revision"
        )
    if "git_revision" in environment_object:
        environment_revision = _require_revision(
            environment_object.get("git_revision"),
            f"{location}.harness_identity.environment.git_revision",
        )
        if environment_revision != revision:
            raise ClaimInputError(
                f"{location}.harness_identity environment and harness revisions differ"
            )
    dirty = identity.get("git_worktree_dirty")
    if dirty is None:
        dirty = environment_object.get("git_worktree_dirty")
    if dirty is not False:
        raise ClaimInputError(f"{location}.harness_identity.git_worktree_dirty must be false")
    if "git_worktree_dirty" in environment_object and environment_object.get("git_worktree_dirty") is not False:
        raise ClaimInputError(
            f"{location}.harness_identity.environment.git_worktree_dirty must be false"
        )
    if "leg" in identity and identity.get("leg") != label:
        raise ClaimInputError(f"{location}.harness_identity.leg does not match {label}")
    if "variant" in identity and identity.get("variant") != expected_variant:
        raise ClaimInputError(
            f"{location}.harness_identity.variant does not match {expected_variant}"
        )

    tool = _require_object(identity.get("tool"), f"{location}.harness_identity.tool")
    for field in ("name", "version", "profile"):
        _require_string(tool.get(field), f"{location}.harness_identity.tool.{field}")

    heaptrack = _require_object(item.get("heaptrack"), f"{location}.heaptrack")
    heap_identity = _require_object(
        heaptrack.get("harness_identity"),
        f"{location}.heaptrack.harness_identity",
    )
    if heap_identity.get("status") != "validated":
        raise ClaimInputError(
            f"{location}.heaptrack.harness_identity.status must be 'validated'"
        )
    harness_digest = _require_sha(
        heap_identity.get("sha256"),
        f"{location}.heaptrack.harness_identity.sha256",
    )
    return binary, tool, harness_digest


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
    binaries: list[dict[str, Any]] = []
    harness_tools: list[dict[str, Any]] = []
    harness_digests: list[str] = []
    for index, leg in enumerate(legs):
        item = _require_object(leg, f"{evidence['id']}.legs[{index}]")
        label = RESOURCE_LEGS[index]
        variant = RESOURCE_VARIANTS[label]
        binary, tool, harness_digest = _resource_leg_identity(
            item,
            label=label,
            expected_revision=(control_revision if variant == "control" else candidate_revision),
            expected_variant=variant,
            location=f"{evidence['id']}.legs[{index}]",
        )
        binaries.append(binary)
        harness_tools.append(tool)
        harness_digests.append(harness_digest)
    declared_order = validation.get("leg_order")
    if declared_order is not None and declared_order != list(RESOURCE_LEGS):
        raise ClaimInputError(f"resource report {evidence['id']} validation leg order mismatch")
    for field, index in (("control_binary_sha256", 0), ("candidate_binary_sha256", 1)):
        declared_binary = validation.get(field)
        if declared_binary is not None:
            declared_binary = _require_sha(
                declared_binary, f"{evidence['id']}.validation.{field}"
            )
            if declared_binary != binaries[index]["binary_sha256"]:
                raise ClaimInputError(
                    f"resource report {evidence['id']} validation {field} does not match its leg"
                )
    declared_harnesses = validation.get("harness_identities")
    if declared_harnesses is not None:
        declared_harnesses = _require_list(
            declared_harnesses, f"{evidence['id']}.validation.harness_identities"
        )
        if len(declared_harnesses) != len(RESOURCE_LEGS):
            raise ClaimInputError(
                f"resource report {evidence['id']} validation harness identity count mismatch"
            )
        for index, raw_identity in enumerate(declared_harnesses):
            declared_identity = _require_object(
                raw_identity,
                f"{evidence['id']}.validation.harness_identities[{index}]",
            )
            label = RESOURCE_LEGS[index]
            variant = RESOURCE_VARIANTS[label]
            if declared_identity.get("leg") != label or declared_identity.get("variant") != variant:
                raise ClaimInputError(
                    f"resource report {evidence['id']} validation harness identity {label} is rebound"
                )
            declared_revision = _require_revision(
                declared_identity.get("git_revision"),
                f"{evidence['id']}.validation.harness_identities[{index}].git_revision",
            )
            expected_revision = control_revision if variant == "control" else candidate_revision
            if declared_revision != expected_revision:
                raise ClaimInputError(
                    f"resource report {evidence['id']} validation harness identity {label} revision mismatch"
                )
            declared_tool = _require_object(
                declared_identity.get("tool"),
                f"{evidence['id']}.validation.harness_identities[{index}].tool",
            )
            if declared_tool != harness_tools[index]:
                raise ClaimInputError(
                    f"resource report {evidence['id']} validation harness tool identity {label} differs"
                )
    if binaries[0] != binaries[3] or binaries[1] != binaries[2] or binaries[0] == binaries[1]:
        raise ClaimInputError(f"resource report {evidence['id']} binary identities are inconsistent")
    if harness_digests[0] != harness_digests[3] or harness_digests[1] != harness_digests[2]:
        raise ClaimInputError(f"resource report {evidence['id']} harness identities drift within a variant")
    if harness_digests[0] == harness_digests[1]:
        raise ClaimInputError(f"resource report {evidence['id']} control and candidate harness identities are identical")
    if harness_tools[0] != harness_tools[3] or harness_tools[1] != harness_tools[2]:
        raise ClaimInputError(f"resource report {evidence['id']} harness tool profiles drift within a variant")
    if any(tool != harness_tools[0] for tool in harness_tools[1:]):
        raise ClaimInputError(f"resource report {evidence['id']} harness tool profiles differ across ABBA legs")
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
        _verify_resource_metric(
            metric,
            item,
            legs,
            max_regression=max_regression,
            evidence_id=evidence["id"],
        )
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
