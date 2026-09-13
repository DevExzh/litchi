#!/usr/bin/env python3
"""Consume the frozen 0553 XLSX campaign evidence and make one decision.

This is an outcome consumer.  It does not build, capture, test, modify source,
or manufacture a pending/pass document.  An ordinary invocation reads the
0553 verifier and the canonical metrics, guard/cap, profile, review, quality,
and source-custody artifacts.  It writes ``decision.json`` only after every
required lane is complete.  ``--schema`` prints the contract without reading
campaign evidence.

The conditional profile lane is enabled only by the complete pilot:
``main.all_frozen_main_gates_pass`` together with the guard and cap admission
booleans.  The allocator lane remains byte evidence; its elapsed samples are
never promoted to latency evidence.
"""

from __future__ import annotations

import argparse
import datetime as _datetime
import hashlib
import importlib.util
import json
import re
import sys
from pathlib import Path
from typing import Any


sys.dont_write_bytecode = True

HERE = Path(__file__).resolve().parent
VERIFY_PATH = HERE / "verify.py"
DECISION_PATH = HERE / "decision.json"
PROFILE_DECISION_PATH = HERE / "profile-decision.json"
QUALITY_PATH = HERE / "quality.json"
REVIEW_PATH = HERE / "adverse-review.json"
SEAL_PATH = HERE / "SHA256SUMS"

SCHEMA = "xlsx_0553_decision_v1"
METRICS_SCHEMA = "xlsx_multisource_edit_metrics_0553_v1"
GUARD_SCHEMA = "litchi.xlsx.guard-cap-analysis.v1"
PROFILE_DECISION_SCHEMA = "xlsx_0553_profile_decision_v1"
REVIEW_SCHEMA = "xlsx_0553_adverse_review_v1"
QUALITY_SCHEMA = "xlsx_0553_quality_v1"
SCOPE = (
    "Matched commit-local compact source-cell proof experiment for source-backed "
    "XLSX MultiSourceEdit; OLE2/OOXML first, ODF deferred, iWork excluded"
)

PLAN_SHA256 = "17d810a24912065fde8c71de6109be983f7c5b3bdcea5fd6584f4725c6499ea4"
RUN_SHA256 = "f1398fcc87dc7bccfc05930276e26a4a8a21106480613238eab49310ae1c23f0"
CAPTURE_SHA256 = "21b0153923e87a331906552457a5f77bb444686d4ce9c9c99e3efcb2dbe5d2c1"
ANALYSIS_INPUTS_SHA256 = (
    "2d95091698893cf402ce85003f407e0352a80b3fd70fa5c4c5e3110f8332a38a"
)

STAGES = ("baseline", "candidate")
SHAPES = ("medium", "dense-sparse", "noncompact", "vendor-extension")

REVIEW_SOURCES = {
    "metrics_adverse": "metrics-analysis.json:comparisons.adverse",
    "metrics_drift": "metrics-analysis.json:repeat_drift_over_five_percent",
    "guard_adverse": "guard-cap-analysis.json:comparison.guard.adverse_flags_over_five_percent",
    "guard_drift": "guard-cap-analysis.json:comparison.guard.same_build_drift_over_five_percent",
    "cap_adverse": "guard-cap-analysis.json:comparison.cap.adverse_flags_over_five_percent",
    "cap_drift": "guard-cap-analysis.json:comparison.cap.same_build_drift_over_five_percent",
}

SHA256_RE = re.compile(r"^[0-9a-f]{64}$")


class DecisionError(ValueError):
    """Malformed, contradictory, or out-of-scope evidence."""


class IncompleteDecision(DecisionError):
    """An evidence lane required for a decision has not arrived."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise DecisionError(message)


def need(path: Path, label: str, *, directory: bool = False) -> Path:
    if not path.exists() or path.is_symlink():
        raise IncompleteDecision(f"{label} is missing or is a symlink")
    if directory:
        require(path.is_dir(), f"{label} is not a directory")
    else:
        require(path.is_file(), f"{label} is not a regular file")
    return path


def read_json(path: Path, label: str) -> Any:
    need(path, label)
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise DecisionError(f"cannot read {label}: {error}") from error


def digest(path: Path, label: str | None = None) -> str:
    label = label or path.as_posix()
    need(path, label)
    value = hashlib.sha256()
    try:
        with path.open("rb") as stream:
            for block in iter(lambda: stream.read(1024 * 1024), b""):
                value.update(block)
    except OSError as error:
        raise DecisionError(f"cannot hash {label}: {error}") from error
    return value.hexdigest()


def hash_value(value: Any, label: str) -> str:
    require(isinstance(value, str) and SHA256_RE.fullmatch(value) is not None,
            f"{label} is not a lowercase SHA-256 digest")
    return value


def nonempty(value: Any, label: str) -> str:
    require(isinstance(value, str) and value.strip(), f"{label} is empty")
    return value


def json_default(value: Any) -> str:
    """Encode only verifier timestamps, retaining timezone and precision.

    ``validate_profiles`` returns its internal receipt interval endpoints as
    ``datetime`` objects when the conditional profile lane is required.  The
    decision keeps those rows as evidence, so the output boundary must encode
    them instead of dropping or replacing them.  Every other non-JSON object
    is rejected deliberately.
    """

    if isinstance(value, _datetime.datetime):
        if value.tzinfo is None or value.utcoffset() is None:
            raise TypeError("datetime evidence must be timezone-aware")
        return value.isoformat()
    raise TypeError(
        f"unsupported decision evidence object: {type(value).__name__}"
    )


def deterministic_json(value: Any) -> bytes:
    """Return the sole deterministic JSON representation used by this script."""

    return (
        json.dumps(value, indent=2, sort_keys=True, allow_nan=False,
                   default=json_default) + "\n"
    ).encode("utf-8")


def relative_to_here(path: Path) -> str:
    try:
        return path.resolve().relative_to(HERE.resolve()).as_posix()
    except ValueError as error:
        raise DecisionError(f"path escapes the 0553 evidence bundle: {path}") from error


def verify_module() -> Any:
    """Load the bundle verifier without executing its command-line interface."""

    need(VERIFY_PATH, "0553 verify.py")
    sys.dont_write_bytecode = True
    spec = importlib.util.spec_from_file_location(
        "xlsx_0553_verify_for_decision", VERIFY_PATH
    )
    require(spec is not None and spec.loader is not None,
            "cannot load the 0553 verifier")
    module = importlib.util.module_from_spec(spec)
    try:
        spec.loader.exec_module(module)
    except (OSError, ImportError, SyntaxError, ValueError) as error:
        raise DecisionError(f"cannot load the 0553 verifier: {error}") from error
    return module


def call(verify: Any, name: str, *args: Any, **kwargs: Any) -> Any:
    function = getattr(verify, name, None)
    if function is None or not callable(function):
        raise IncompleteDecision(f"verifier component {name} is unavailable")
    try:
        return function(*args, **kwargs)
    except IncompleteDecision:
        raise
    except DecisionError:
        raise
    except Exception as error:  # verifier errors must fail closed
        # The verifier owns its exception classes.  Preserve the useful
        # incomplete/evidence distinction without importing an implementation
        # module's private classes into this consumer.
        if error.__class__.__name__.lower().startswith("incomplete"):
            raise IncompleteDecision(
                f"verifier component {name} is incomplete: {error}"
            ) from error
        raise DecisionError(f"verifier component {name} failed: {error}") from error


def report_document(verify: Any, result: dict[str, Any], label: str,
                    expected_schema: str) -> tuple[dict[str, Any], Path, str]:
    require(isinstance(result, dict) and result.get("status") == "pass",
            f"{label} verifier result is not a completed pass")
    report = result.get("report")
    require(isinstance(report, dict), f"{label} report reference is missing")
    report_name = report.get("path")
    require(isinstance(report_name, str) and report_name,
            f"{label} report path is missing")
    candidate = Path(report_name)
    path = candidate if candidate.is_absolute() else HERE / candidate
    relative_to_here(path)
    document = read_json(path, f"{label} canonical report")
    require(isinstance(document, dict)
            and document.get("schema") == expected_schema,
            f"{label} canonical report schema differs")
    observed = digest(path, f"{label} canonical report")
    require(report.get("sha256") == observed,
            f"{label} canonical report digest differs")
    require(document.get("status") == "pass",
            f"{label} canonical report is not a completed pass")
    return document, path, observed


def main_gate_rows(metrics: dict[str, Any]) -> dict[str, Any]:
    """Require the complete frozen main gate inventory and every check row."""

    gates = metrics.get("main_gates")
    require(isinstance(gates, dict), "main metrics gates are missing")
    expected = {
        "primary_one_percent", "one_cell_latency", "workflow_memory",
        "allocation", "correctness_identity", "all_frozen_main_gates_pass",
        "external_controls_required",
    }
    require(set(gates) == expected, "main metrics gate inventory differs")
    for name in ("primary_one_percent", "one_cell_latency", "workflow_memory",
                 "allocation"):
        group = gates[name]
        require(isinstance(group, dict)
                and isinstance(group.get("pass"), bool)
                and isinstance(group.get("checks"), list)
                and group["checks"],
                f"main metrics {name} checks are missing")
        if "check_count" in group:
            require(isinstance(group["check_count"], int)
                    and not isinstance(group["check_count"], bool)
                    and group["check_count"] == len(group["checks"]),
                    f"main metrics {name}.check_count differs")
        for index, row in enumerate(group["checks"]):
            require(isinstance(row, dict) and isinstance(row.get("pass"), bool),
                    f"main metrics {name}.checks[{index}] is malformed")
    correctness = gates["correctness_identity"]
    require(isinstance(correctness, dict)
            and isinstance(correctness.get("pass"), bool)
            and isinstance(correctness.get("identity_equal"), bool)
            and isinstance(correctness.get("exact_checks"), list)
            and correctness.get("exact_checks"),
            "main correctness identity checks are missing")
    for index, row in enumerate(correctness["exact_checks"]):
        require(isinstance(row, dict) and isinstance(row.get("equal"), bool),
                f"main correctness exact check {index} is malformed")
    require(isinstance(gates["all_frozen_main_gates_pass"], bool),
            "main aggregate gate is not boolean")
    require(gates["external_controls_required"] == {
        "status": "pending", "validated_here": False,
        "required": ["guard", "cap", "quality", "profile"],
        "reason": "main metrics analyzer does not own guard, cap, quality, or profile evidence",
    }, "main external-control gate differs")
    return gates


def extract(value: Any, path: str) -> Any:
    current = value
    for part in path.split("."):
        require(isinstance(current, dict) and part in current,
                f"canonical report field is missing: {path}")
        current = current[part]
    return current


def gate_rows_pass(value: Any, label: str) -> None:
    """Require each supplemental check row to expose a boolean result."""

    require(isinstance(value, dict), f"{label} gate object is missing")
    gates = value.get("gates")
    if isinstance(gates, list):
        groups = [("gates", gates)]
    else:
        require(isinstance(gates, dict), f"{label} gate rows are missing")
        groups = list(gates.items())
    require(groups, f"{label} has no gate groups")
    for name, rows in groups:
        require(isinstance(rows, list) and rows,
                f"{label}.{name} has no explicit checks")
        for index, row in enumerate(rows):
            require(isinstance(row, dict) and isinstance(row.get("passed"), bool),
                    f"{label}.{name}[{index}] has no explicit passed flag")


def guard_cap_evidence(guards: dict[str, Any],
                       result: dict[str, Any]) -> tuple[bool, bool]:
    comparison = extract(guards, "comparison")
    guard = comparison.get("guard")
    cap = comparison.get("cap")
    require(isinstance(guard, dict) and isinstance(cap, dict),
            "guard/cap comparison inventory differs")
    guard_gate = result.get("guard_admission_passed")
    cap_gate = result.get("cap_admission_passed")
    if not isinstance(guard_gate, bool):
        guard_gate = guard.get("admission_passed")
    if not isinstance(cap_gate, bool):
        cap_gate = cap.get("admission_passed")
    require(isinstance(guard_gate, bool) and isinstance(cap_gate, bool),
            "guard/cap admission booleans are missing")
    require(guard.get("admission_passed") is guard_gate
            and cap.get("admission_passed") is cap_gate,
            "guard/cap admission aliases differ")
    require(guards.get("admission_status") ==
            ("pass" if guard_gate and cap_gate else "reject"),
            "guard/cap admission status differs from both gates")
    gate_rows_pass(guard, "guard")
    gate_rows_pass(cap, "cap")
    for path in (
        "comparison.guard.adverse_flags_over_five_percent",
        "comparison.guard.same_build_drift_over_five_percent",
        "comparison.cap.adverse_flags_over_five_percent",
        "comparison.cap.same_build_drift_over_five_percent",
    ):
        require(isinstance(extract(guards, path), list),
                f"{path} is not an array")
    return guard_gate, cap_gate


def reviewed_rows(raw: list[Any], checked: Any, source: str) -> list[dict[str, Any]]:
    """Match every raw adverse/drift row to exactly one annotated row."""

    require(isinstance(checked, list) and len(checked) == len(raw),
            f"{source} review length differs")
    remaining = list(checked)
    normalized: list[dict[str, Any]] = []
    for index, original in enumerate(raw):
        require(isinstance(original, dict), f"{source} raw row {index} is malformed")
        matches = [
            (position, row) for position, row in enumerate(remaining)
            if isinstance(row, dict) and row.get("original") == original
        ]
        require(len(matches) == 1,
                f"{source} raw row {index} is not individually reviewed")
        position, row = matches[0]
        if "source" in row:
            require(row["source"] == source,
                    f"{source} row {index} source label differs")
        for field in ("id", "classification", "interpretation", "disposition"):
            nonempty(row.get(field), f"{source} row {index}.{field}")
        normalized.append(row)
        remaining.pop(position)
    require(not remaining, f"{source} review has extra rows")
    return normalized


def validate_adverse_review(metrics: dict[str, Any], metrics_sha: str,
                            guards: dict[str, Any], guards_sha: str) -> dict[str, Any]:
    path = need(REVIEW_PATH, "adverse-review.json")
    review = read_json(path, "adverse-review.json")
    require(isinstance(review, dict)
            and review.get("schema") == REVIEW_SCHEMA
            and review.get("status") == "complete"
            and review.get("complete") is True
            and review.get("all_diagnostic_rows_retained") is True,
            "adverse review completion envelope differs")
    require(review.get("comparison_sha256") == metrics_sha
            and review.get("guard_analysis_sha256") == guards_sha,
            "adverse review analyzer digest binding differs")
    require(isinstance(review.get("adoption_allowed"), bool),
            "adverse review adoption result is missing")
    raw = {
        "metrics_adverse": extract(metrics, "comparisons.adverse"),
        "metrics_drift": extract(metrics, "repeat_drift_over_five_percent"),
        "guard_adverse": extract(
            guards, "comparison.guard.adverse_flags_over_five_percent"),
        "guard_drift": extract(
            guards, "comparison.guard.same_build_drift_over_five_percent"),
        "cap_adverse": extract(
            guards, "comparison.cap.adverse_flags_over_five_percent"),
        "cap_drift": extract(
            guards, "comparison.cap.same_build_drift_over_five_percent"),
    }
    require(all(isinstance(rows, list) for rows in raw.values()),
            "adverse or drift source is not an array")
    groups = review.get("groups")
    retained: dict[str, list[dict[str, Any]]] = {}
    if groups is not None:
        require(isinstance(groups, list) and len(groups) == len(REVIEW_SOURCES),
                "adverse review group inventory differs")
        by_source: dict[str, Any] = {}
        for index, group in enumerate(groups):
            require(isinstance(group, dict)
                    and set(group) == {"source", "rows"},
                    f"adverse review group {index} is malformed")
            source = nonempty(group["source"],
                             f"adverse review group {index}.source")
            require(source not in by_source,
                    f"adverse review repeats source: {source}")
            require(isinstance(group["rows"], list),
                    f"adverse review group {index}.rows is not an array")
            by_source[source] = group["rows"]
        require(set(by_source) == set(REVIEW_SOURCES.values()),
                "adverse review source inventory differs")
        for key, source in REVIEW_SOURCES.items():
            retained[key] = reviewed_rows(raw[key], by_source[source], source)
    else:
        aliases = {
            "metrics_adverse": ("metrics_adverse", "matched", "matched_adverse_flags_over_five_percent"),
            "metrics_drift": ("metrics_drift", "same_build", "same_build_variations_over_five_percent"),
            "guard_adverse": ("guard_adverse", "guard_adverse_flags"),
            "guard_drift": ("guard_drift", "guard_same_build_drift_flags"),
            "cap_adverse": ("cap_adverse", "cap_adverse_flags"),
            "cap_drift": ("cap_drift", "cap_same_build_drift_flags"),
        }
        for key, names in aliases.items():
            present = [name for name in names if name in review]
            require(len(present) == 1,
                    f"adverse review named array is missing or duplicated: {key}")
            retained[key] = reviewed_rows(raw[key], review[present[0]],
                                          REVIEW_SOURCES[key])
    expected_count = sum(len(rows) for rows in raw.values())
    counts = review.get("counts")
    if counts is not None:
        require(isinstance(counts, dict)
                and counts.get("reviewed_flags") == expected_count,
                "adverse review count does not cover every source row")
    coverage = review.get("source_coverage")
    if coverage is not None:
        require(isinstance(coverage, list),
                "adverse review source coverage is malformed")
        expected_coverage = [
            {"source": REVIEW_SOURCES[key], "count": len(raw[key])}
            for key in REVIEW_SOURCES
        ]
        require(sorted(coverage, key=lambda row: (row.get("source", ""), row.get("count", -1)))
                == sorted(expected_coverage, key=lambda row: (row["source"], row["count"])),
                "adverse review source coverage differs")
    return {
        "status": "pass", "path": relative_to_here(path),
        "sha256": digest(path), "schema": REVIEW_SCHEMA, "complete": True,
        "adoption_allowed": review["adoption_allowed"],
        "counts": {key: len(rows) for key, rows in raw.items()},
        "reviewed_flags": expected_count,
        "groups": [
            {"source": REVIEW_SOURCES[key], "rows": retained[key]}
            for key in REVIEW_SOURCES
        ],
        "document": review,
    }


def stage_manifest(verify: Any, stage: str) -> tuple[dict[str, str], str]:
    function = getattr(verify, "stage_manifest", None)
    if function is None or not callable(function):
        raise IncompleteDecision("verifier stage_manifest is unavailable")
    try:
        value = function(stage)
    except Exception as error:
        raise DecisionError(f"verifier stage_manifest({stage}) failed: {error}") from error
    require(isinstance(value, tuple) and len(value) == 2,
            f"verifier stage_manifest({stage}) result is malformed")
    manifest, observed = value
    require(isinstance(manifest, dict) and manifest,
            f"{stage} source manifest is empty")
    observed = hash_value(observed, f"{stage} source manifest digest")
    path = HERE / stage / "source-manifest.json"
    require(digest(path, f"{stage}/source-manifest.json") == observed,
            f"{stage} source manifest digest is unstable")
    return dict(manifest), observed


def retained_paths(value: Any) -> tuple[set[str], dict[str, str]]:
    """Read explicit retained-test names and hashes from verifier custody."""

    if isinstance(value, dict):
        for key in ("retained_tests", "retained_test_binding", "retained_tests_binding"):
            if key in value:
                return retained_paths(value[key])
        paths = value.get("paths", value.get("files", value.get("tests")))
        if isinstance(paths, dict):
            names: set[str] = set()
            hashes: dict[str, str] = {}
            for name, item in paths.items():
                name = nonempty(name, "retained test path")
                names.add(name)
                hashes[name] = hash_value(item, f"retained test {name}")
            return names, hashes
        if isinstance(paths, list):
            names = set()
            hashes: dict[str, str] = {}
            for item in paths:
                if isinstance(item, str):
                    names.add(item)
                elif isinstance(item, dict):
                    name = item.get("path", item.get("name"))
                    names.add(nonempty(name, "retained test path"))
                    if "sha256" in item:
                        hashes[name] = hash_value(item["sha256"],
                                                  f"retained test {name}")
                else:
                    raise DecisionError("retained test row is malformed")
            return names, hashes
    if isinstance(value, list):
        return retained_paths({"paths": value})
    raise DecisionError("explicit retained-test binding is malformed")


def validate_final_source(verify: Any, adoption: bool,
                          candidate_manifest: dict[str, str],
                          candidate_sha: str) -> dict[str, Any]:
    """Bind final checkout to candidate acceptance or exact baseline custody."""

    function = getattr(verify, "validate_final_source", None)
    if function is None or not callable(function):
        raise IncompleteDecision("verifier component validate_final_source is unavailable")
    disposition = "accepted" if adoption else "rejected"
    try:
        value = function(disposition, candidate_manifest, candidate_sha)
    except IncompleteDecision:
        raise
    except Exception as error:
        raise DecisionError(f"verifier final-source validation failed: {error}") from error
    require(isinstance(value, dict) and value.get("status") == "pass",
            "final source custody is not a completed pass")
    final_manifest, final_sha = stage_manifest(verify, "final")
    require(value.get("manifest_sha256", final_sha) == final_sha,
            "final source manifest digest differs from verifier")
    baseline_manifest, baseline_sha = stage_manifest(verify, "baseline")
    if adoption:
        require(final_manifest == candidate_manifest and final_sha == candidate_sha,
                "accepted final source is not exactly the candidate")
        selected = "candidate"
    else:
        if final_manifest == baseline_manifest:
            require(final_sha == baseline_sha,
                    "restored baseline source manifest digest differs")
            selected = "baseline"
        else:
            # Any retained tests must be explicit custody data from the
            # verifier.  A changed manifest with no named/hash-bound retained
            # set cannot be treated as a harmless restoration.
            binding = next((value[key] for key in (
                "retained_tests", "retained_test_binding", "retained_tests_binding"
            ) if key in value), None)
            require(binding is not None,
                    "rejected final source has unbound retained files")
            names, hashes = retained_paths(binding)
            require(names, "rejected retained-test binding is empty")
            require(set(hashes) == names,
                    "retained-test binding must hash every retained path")
            changed = {
                name for name in set(baseline_manifest) | set(final_manifest)
                if baseline_manifest.get(name) != final_manifest.get(name)
            }
            require(changed <= names,
                    "rejected final source changes files outside retained tests")
            require(names <= set(final_manifest),
                    "retained-test binding names absent from final source")
            for name, value_hash in hashes.items():
                require(final_manifest.get(name) == value_hash,
                        f"retained test hash differs: {name}")
            selected = "baseline-plus-retained-tests"
    live_function = getattr(verify, "current_source_manifest", None)
    require(callable(live_function), "verifier current_source_manifest is unavailable")
    try:
        live = live_function()
    except Exception as error:
        raise DecisionError(f"live source custody validation failed: {error}") from error
    require(live == final_manifest, "live source does not equal final source custody")
    patch = need(HERE / "final" / "source.patch", "final/source.patch")
    result = {
        "status": "pass", "stage": "final", "manifest_sha256": final_sha,
        "manifest_entries": len(final_manifest), "selected": selected,
        "patch_sha256": digest(patch, "final/source.patch"),
        "candidate_manifest_sha256": candidate_sha,
    }
    if selected == "baseline-plus-retained-tests":
        result["retained_tests"] = binding
    return result


def validate_quality_final(verify: Any, final_sha: str) -> dict[str, Any]:
    quality_result = call(verify, "validate_quality")
    require(isinstance(quality_result, dict)
            and quality_result.get("status") == "pass",
            "quality verifier did not complete a pass")
    document = read_json(QUALITY_PATH, "quality.json")
    require(isinstance(document, dict)
            and document.get("schema") == QUALITY_SCHEMA
            and document.get("status") == "pass"
            and document.get("source_stage") == "final"
            and document.get("source_manifest_sha256") == final_sha,
            "quality result is not bound to final source custody")
    completed = nonempty(document.get("completed_utc"),
                         "quality.completed_utc")
    selected = quality_result.get("selected_attempt")
    if selected is None:
        selected = quality_result.get("attempt")
    nonempty(selected, "quality selected attempt")
    if "canonical_sha256" in quality_result:
        require(quality_result["canonical_sha256"] == digest(QUALITY_PATH),
                "quality verifier digest differs")
    return {
        "status": "pass", "path": relative_to_here(QUALITY_PATH),
        "sha256": digest(QUALITY_PATH), "schema": QUALITY_SCHEMA,
        "selected_attempt": selected, "completed_utc": completed,
        "source_stage": "final", "source_manifest_sha256": final_sha,
        "commands": document.get("commands"), "document": document,
    }


def validate_profile_decision(verify: Any, metrics_sha: str,
                              profiles: dict[str, Any],
                              pilot_gates: dict[str, bool]) -> dict[str, Any]:
    """Bind the explicit profile pass or the explicit skipped-pilot record."""

    path = need(PROFILE_DECISION_PATH, "profile-decision.json")
    value = read_json(path, "profile-decision.json")
    expected_keys = {
        "schema", "status", "scope", "pilot_passed", "profile_required",
        "profile_gate_passed", "main_analysis_sha256", "pilot_gates",
        "profile_rows", "reason",
    }
    require(isinstance(value, dict) and set(value) == expected_keys,
            "profile-decision envelope differs")
    require(value["schema"] == PROFILE_DECISION_SCHEMA,
            "profile-decision schema differs")
    pilot_passed = all(pilot_gates.values())
    require(value["pilot_passed"] is pilot_passed
            and value["main_analysis_sha256"] == metrics_sha,
            "profile-decision pilot/report binding differs")
    hash_value(value["main_analysis_sha256"],
               "profile-decision.main_analysis_sha256")
    require(isinstance(value["pilot_gates"], dict)
            and set(value["pilot_gates"]) == {"main", "guard", "cap"}
            and all(isinstance(item, bool)
                    for item in value["pilot_gates"].values())
            and value["pilot_gates"] == pilot_gates,
            "profile-decision pilot gates differ")
    nonempty(value["scope"], "profile-decision.scope")
    nonempty(value["reason"], "profile-decision.reason")
    require(isinstance(value["profile_rows"], list),
            "profile-decision profile_rows is not an array")
    profile_required = profiles.get("required")
    profile_gate = profiles.get("gate_passed")
    require(isinstance(profile_required, bool)
            and isinstance(profile_gate, bool),
            "profile verifier required/gate fields are missing")
    require(value["profile_required"] is profile_required
            and value["profile_gate_passed"] is profile_gate,
            "profile-decision fields differ from profile verifier")
    decision_ref = profiles.get("decision")
    if decision_ref is not None:
        require(decision_ref == relative_to_here(path),
                "profile verifier selected a different decision record")
    if not pilot_passed:
        require(value["status"] == "skipped"
                and value["profile_required"] is False
                and value["profile_gate_passed"] is True
                and value["profile_rows"] == [],
                "failed pilot lacks an explicit vacuous profile decision")
        for stage in STAGES:
            require(not list((HERE / stage).glob("profile-*.receipt.json")),
                    f"{stage} profile receipts exist after a skipped pilot")
    else:
        require(value["status"] in {"pass", "failed", "reject", "rejected"}
                and value["profile_required"] is True
                and value["profile_rows"],
                "required profile lane is incomplete")
        require(all(isinstance(row, dict)
                    and isinstance(row.get("passed"), bool)
                    for row in value["profile_rows"]),
                "profile-decision rows lack explicit passed flags")
    return {
        "status": "pass", "path": relative_to_here(path),
        "sha256": digest(path), "schema": PROFILE_DECISION_SCHEMA,
        "pilot_passed": pilot_passed, "required": profile_required,
        "gate_passed": profile_gate, "pilot_gates": dict(pilot_gates),
        "profile_rows": value["profile_rows"], "document": value,
    }


def validate_report_bindings(metrics: dict[str, Any], guards: dict[str, Any]) -> None:
    """Reject stale canonical reports even if an older verifier is permissive."""

    require(digest(HERE / "plan.json") == PLAN_SHA256,
            "plan is not the frozen 0553 plan")
    require(digest(HERE / "run.py") == RUN_SHA256,
            "run.py is not the frozen 0553 driver")
    require(digest(HERE / "capture.py") == CAPTURE_SHA256,
            "capture.py is not the frozen 0553 driver")
    require(metrics.get("plan_sha256") == PLAN_SHA256
            and metrics.get("run_sha256") == RUN_SHA256
            and metrics.get("capture_sha256") == CAPTURE_SHA256,
            "main report driver binding differs")
    require(guards.get("plan_sha256") == PLAN_SHA256
            and guards.get("run_sha256") == RUN_SHA256
            and guards.get("capture_sha256") == CAPTURE_SHA256,
            "guard/cap report driver binding differs")
    if "analysis_inputs_sha256" in metrics:
        require(metrics["analysis_inputs_sha256"] == ANALYSIS_INPUTS_SHA256,
                "main analysis-input binding differs")
    if "analysis_inputs_sha256" in guards:
        require(guards["analysis_inputs_sha256"] == ANALYSIS_INPUTS_SHA256,
                "guard analysis-input binding differs")


def evaluate() -> dict[str, Any]:
    verify = verify_module()

    metrics_result = call(verify, "validate_metrics_analysis")
    guards_result = call(verify, "validate_guard_cap_analysis")
    metrics, metrics_path, metrics_sha = report_document(
        verify, metrics_result, "main metrics", METRICS_SCHEMA
    )
    guards, guards_path, guards_sha = report_document(
        verify, guards_result, "guard/cap", GUARD_SCHEMA
    )
    validate_report_bindings(metrics, guards)
    main_gates = main_gate_rows(metrics)
    guard_gate, cap_gate = guard_cap_evidence(guards, guards_result)
    main_gate = bool(main_gates["all_frozen_main_gates_pass"])
    pilot_gates = {"main": main_gate, "guard": guard_gate, "cap": cap_gate}
    pilot_passed = all(pilot_gates.values())

    # The complete pilot, rather than a convenient native/memory subset, is
    # the only condition that can make exact commit profiles mandatory.
    profiles = call(verify, "validate_profiles", pilot_expected=pilot_passed)
    require(isinstance(profiles, dict)
            and profiles.get("pilot_passed") is pilot_passed,
            "profile verifier pilot result differs")
    profile_gate = profiles.get("gate_passed")
    require(isinstance(profile_gate, bool), "profile gate is missing")
    if pilot_passed:
        require(profiles.get("required") is True,
                "exact commit profile lane was not required")
    else:
        require(profiles.get("required") is False and profile_gate is True,
                "skipped profile lane is not explicitly vacuous")
    profile_decision = validate_profile_decision(
        verify, metrics_sha, profiles, pilot_gates
    )

    review = validate_adverse_review(metrics, metrics_sha, guards, guards_sha)
    review_gate = bool(review["adoption_allowed"])
    adoption = bool(main_gate and guard_gate and cap_gate
                    and profile_gate and review_gate)

    candidate_manifest, candidate_sha = stage_manifest(verify, "candidate")
    final_source = validate_final_source(
        verify, adoption, candidate_manifest, candidate_sha
    )
    quality = validate_quality_final(verify, final_source["manifest_sha256"])
    quality_gate = quality["status"] == "pass"
    require(quality_gate, "final quality gate did not pass")
    if adoption:
        require(final_source["selected"] == "candidate",
                "accepted disposition does not select the candidate source")
    else:
        require(final_source["selected"] in {
            "baseline", "baseline-plus-retained-tests"
        }, "rejected disposition does not select a restored source")

    gate_map = {
        "main": main_gate, "guard": guard_gate, "cap": cap_gate,
        "profile": profile_gate,
        "profile_required": bool(profiles.get("required")),
        "quality": quality_gate, "adverse_review": review_gate,
    }
    return {
        "schema": SCHEMA, "status": "pass", "scope": SCOPE,
        "disposition": "accepted" if adoption else "rejected",
        "adoption_allowed": adoption,
        "observed_utc": quality["completed_utc"],
        "gates": gate_map,
        "main_gate": main_gate, "native_primary_gate": main_gate,
        "guard_gate": guard_gate, "cap_gate": cap_gate,
        "profile_gate": profile_gate, "quality_gate": quality_gate,
        "adverse_review_gate": review_gate,
        "main_gates": main_gates,
        "pilot": {"passed": pilot_passed,
                   "gates": pilot_gates,
                   "profile_required": profiles["required"]},
        "evidence": {
            "metrics": {
                "path": relative_to_here(metrics_path), "sha256": metrics_sha,
                "schema": metrics["schema"], "status": metrics["status"],
                "matched_identity": metrics.get("matched_identity"),
                "comparisons": metrics.get("comparisons"),
                "repeat_drift": metrics.get("repeat_drift"),
                "repeat_drift_over_five_percent": metrics.get(
                    "repeat_drift_over_five_percent"),
                "main_gates": main_gates,
                "allocation": metrics.get("allocation"),
                "field_paths": metrics.get("field_paths"),
                "full_report_preserved": True,
            },
            "guards": {
                "path": relative_to_here(guards_path), "sha256": guards_sha,
                "schema": guards["schema"], "status": guards["status"],
                "comparison": guards.get("comparison"),
                "full_report_preserved": True,
            },
            "profiles": profiles,
            "profile_decision": profile_decision,
            "quality": quality,
            "adverse_review": review,
        },
        "final_source": final_source,
        "source_manifest_sha256": final_source["manifest_sha256"],
        "metrics_analysis_sha256": metrics_sha,
        "main_analysis_sha256": metrics_sha,
        "guard_analysis_sha256": guards_sha,
        "quality_sha256": quality["sha256"],
        "quality_summary_sha256": quality["sha256"],
        "adverse_review_sha256": review["sha256"],
        "profile_analysis_sha256": profiles.get("decision_sha256"),
        "all_metrics_preserved": True,
        "exact_review_sources": list(REVIEW_SOURCES.values()),
    }


def schema_document() -> dict[str, Any]:
    return {
        "schema": SCHEMA, "scope": SCOPE,
        "mode": "outcome-neutral; fail closed; no synthetic evidence",
        "frozen_bindings": {
            "revision": "8aa0c5baf0616d16c79eba0c6c28dc1716338ad6",
            "plan_sha256": PLAN_SHA256, "run_sha256": RUN_SHA256,
            "capture_sha256": CAPTURE_SHA256,
            "analysis_inputs_sha256": ANALYSIS_INPUTS_SHA256,
        },
        "required_evidence": {
            "main": {
                "canonical_report": "metrics-analysis.json",
                "schema": METRICS_SCHEMA,
                "gate": "main_gates.all_frozen_main_gates_pass",
                "identity": "main_gates.correctness_identity",
                "preservation": "retain canonical comparisons, drift, allocation field paths, and full report hash",
            },
            "guard_cap": {
                "canonical_report": "guard-cap-analysis.json",
                "schema": GUARD_SCHEMA,
                "gates": [
                    "comparison.guard.admission_passed",
                    "comparison.cap.admission_passed",
                ],
                "rows": "every guard/cap gate row has an explicit passed boolean",
            },
            "profile": {
                "condition": (
                    "main_gates.all_frozen_main_gates_pass and "
                    "comparison.guard.admission_passed and "
                    "comparison.cap.admission_passed"
                ),
                "canonical": "profile-decision.json",
                "schema": PROFILE_DECISION_SCHEMA,
                "skip": (
                    "when the complete pilot is false: status=skipped, "
                    "profile_required=false, profile_gate_passed=true, profile_rows=[]"
                ),
                "required": "when pilot is true, retain every profile row and its passed boolean",
            },
            "adverse_review": {
                "canonical": "adverse-review.json",
                "schema": REVIEW_SCHEMA,
                "hashes": ["comparison_sha256", "guard_analysis_sha256"],
                "groups": list(REVIEW_SOURCES.values()),
                "row_rule": "each source array is matched one-for-one by an original row plus id/classification/interpretation/disposition",
            },
            "source_custody": {
                "candidate": "verifier candidate stage manifest",
                "accepted": "final manifest exactly equals candidate manifest and digest",
                "rejected": (
                    "final manifest exactly equals baseline, or differs only by an "
                    "explicit verifier-retained test set whose paths and hashes are bound"
                ),
                "live": "verifier current_source_manifest must equal final manifest",
            },
            "quality": {
                "canonical": "quality.json",
                "schema": QUALITY_SCHEMA,
                "binding": "status=pass, source_stage=final, source_manifest_sha256=final manifest",
            },
        },
        "decision_rule": (
            "accepted iff every frozen main, guard, cap, conditional profile, "
            "quality, and adverse-review adoption gate passes; otherwise rejected "
            "only after final source restoration custody and quality pass"
        ),
        "no_claims": [
            "No decision is emitted while required evidence is missing.",
            "Allocator elapsed samples are not latency evidence.",
            "ODF remains deferred until the OLE2/OOXML optimization goal completes; iWork is excluded.",
        ],
    }


def write_identical(path: Path, value: dict[str, Any]) -> None:
    # Use the same strict serializer for exclusive-create and replay.  This
    # is the boundary at which verifier-only datetime objects become stable
    # ISO-8601 strings; unknown objects remain a hard failure.
    encoded = deterministic_json(value)
    if path.exists():
        require(path.is_file() and not path.is_symlink(),
                f"decision output is not a regular file: {path}")
        require(path.read_bytes() == encoded,
                f"existing decision output differs; refusing replacement: {path}")
        return
    if SEAL_PATH.exists() and path.resolve().is_relative_to(HERE.resolve()):
        raise DecisionError("sealed evidence has no retained decision output")
    path.parent.mkdir(parents=True, exist_ok=True)
    try:
        with path.open("xb") as stream:
            stream.write(encoded)
    except FileExistsError:
        require(path.read_bytes() == encoded,
                f"decision output changed during exclusive create: {path}")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--schema", action="store_true",
                        help="print the contract without reading campaign evidence")
    parser.add_argument("--output", type=Path, default=DECISION_PATH,
                        help="decision output; exclusive-create/identical-replay")
    args = parser.parse_args(argv)
    if args.schema:
        print(json.dumps(schema_document(), indent=2, sort_keys=True))
        return 0
    try:
        result = evaluate()
        write_identical(args.output, result)
    except IncompleteDecision as error:
        print(f"0553 decision incomplete: {error}", file=sys.stderr)
        return 2
    except (DecisionError, OSError, KeyError, TypeError, ValueError) as error:
        print(f"0553 decision rejected by evidence checks: {error}", file=sys.stderr)
        return 2
    print(deterministic_json(result).decode("utf-8"), end="")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
