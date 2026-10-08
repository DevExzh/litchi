#!/usr/bin/env python3
"""Derive a fail-closed qualification decision from validated facet facts.

This module is deliberately one layer above the evidence validators.  An
auditor supplies separate facts for completeness, custody, correctness,
readiness, performance, and benefit.  This module does not parse raw packets,
choose thresholds, or accept a precomputed verdict; it only applies the
published precedence rules to those facts.

The command line interface requires both an input facts file and a fresh JSON
output path::

    python3 tools/perf_decision.py \
        --input decision-input.json --json-out decision-output.json

Existing output files are never replaced.  ``--help`` and rejected input do
not create output files.
"""

from __future__ import annotations

import argparse
import json
import os
import re
import sys
from dataclasses import dataclass
from enum import Enum
from pathlib import Path
from typing import Any, Mapping, Optional, Sequence, Tuple


INPUT_SCHEMA = "litchi.performance.qualification_decision_input.v1"
OUTPUT_SCHEMA = "litchi.performance.qualification_decision.v1"
TOOL_NAME = "litchi-perf-decision"
TOOL_VERSION = "0.1.0"

_FACET_NAMES = ("custody", "correctness", "readiness", "performance", "benefit")
_REASON_PATTERN = re.compile(r"^[a-z][a-z0-9_.-]{0,127}$")
_SHA256_PATTERN = re.compile(r"^[0-9a-f]{64}$")


class DecisionInputError(ValueError):
    """The facts document is not a valid decision-input document."""


class DecisionOutputError(OSError):
    """The derived decision could not be written to a fresh destination."""


class FacetStatus(str, Enum):
    PASS = "pass"
    FAIL = "fail"
    MISSING = "missing"
    UNINTERPRETABLE = "uninterpretable"


class CompletenessStatus(str, Enum):
    COMPLETE = "complete"
    MISSING = "missing"


class DecisionCode(str, Enum):
    ADOPT = "adopt"
    REJECT_CANDIDATE = "reject_candidate"
    INCOMPLETE = "incomplete"
    UNINTERPRETABLE = "uninterpretable"


def _validate_reason_codes(reason_codes: Sequence[str], context: str) -> Tuple[str, ...]:
    if type(reason_codes) is not tuple:
        raise TypeError(f"{context}.reason_codes must be a tuple")
    if not reason_codes:
        raise ValueError(f"{context}.reason_codes must not be empty")
    if any(type(code) is not str for code in reason_codes):
        raise TypeError(f"{context}.reason_codes must contain strings")
    if any(not _REASON_PATTERN.fullmatch(code) for code in reason_codes):
        raise ValueError(
            f"{context}.reason_codes must contain lowercase reason codes"
        )
    if len(set(reason_codes)) != len(reason_codes):
        raise ValueError(f"{context}.reason_codes must not contain duplicates")
    return reason_codes


@dataclass(frozen=True)
class Completeness:
    status: CompletenessStatus
    reason_codes: Tuple[str, ...]

    def __post_init__(self) -> None:
        if type(self.status) is not CompletenessStatus:
            raise TypeError("completeness.status must be CompletenessStatus")
        _validate_reason_codes(self.reason_codes, "completeness")

    def as_dict(self) -> dict[str, Any]:
        return {
            "status": self.status.value,
            "reason_codes": list(self.reason_codes),
        }


@dataclass(frozen=True)
class Facet:
    status: FacetStatus
    reason_codes: Tuple[str, ...]

    def __post_init__(self) -> None:
        if type(self.status) is not FacetStatus:
            raise TypeError("facet.status must be FacetStatus")
        _validate_reason_codes(self.reason_codes, "facet")

    def as_dict(self) -> dict[str, Any]:
        return {
            "status": self.status.value,
            "reason_codes": list(self.reason_codes),
        }


@dataclass(frozen=True)
class ProvenanceRef:
    kind: str
    path: str
    sha256: str

    def __post_init__(self) -> None:
        for field_name, value in (
            ("kind", self.kind),
            ("path", self.path),
            ("sha256", self.sha256),
        ):
            if type(value) is not str or not value:
                raise ValueError(f"provenance.{field_name} must be a nonempty string")
        if not _SHA256_PATTERN.fullmatch(self.sha256):
            raise ValueError("provenance.sha256 must be 64 lowercase hex characters")

    def as_dict(self) -> dict[str, str]:
        return {"kind": self.kind, "path": self.path, "sha256": self.sha256}


@dataclass(frozen=True)
class QualificationFacts:
    """Validated facts from an evidence audit, without a derived verdict."""

    candidate: str
    source_revision: str
    completeness: Completeness
    custody: Facet
    correctness: Facet
    readiness: Facet
    performance: Facet
    benefit: Facet
    provenance: Tuple[ProvenanceRef, ...]

    def __post_init__(self) -> None:
        for field_name, value in (
            ("candidate", self.candidate),
            ("source_revision", self.source_revision),
        ):
            if type(value) is not str or not value:
                raise ValueError(f"{field_name} must be a nonempty string")
        if type(self.completeness) is not Completeness:
            raise TypeError("completeness must be Completeness")
        for name in _FACET_NAMES:
            if type(getattr(self, name)) is not Facet:
                raise TypeError(f"{name} must be Facet")
        if type(self.provenance) is not tuple or not self.provenance:
            raise ValueError("provenance must be a nonempty tuple")
        if any(type(item) is not ProvenanceRef for item in self.provenance):
            raise TypeError("provenance must contain ProvenanceRef values")

    def facets(self) -> dict[str, Facet]:
        return {name: getattr(self, name) for name in _FACET_NAMES}


@dataclass(frozen=True)
class DecisionResult:
    facts: QualificationFacts
    decision: DecisionCode
    candidate_adoption: Optional[bool]
    reasons: Tuple[str, ...]

    def __post_init__(self) -> None:
        if type(self.decision) is not DecisionCode:
            raise TypeError("decision must be DecisionCode")
        if type(self.reasons) is not tuple or not self.reasons:
            raise ValueError("decision reasons must not be empty")
        if any(type(reason) is not str or not reason for reason in self.reasons):
            raise TypeError("decision reasons must contain nonempty strings")
        expected: dict[DecisionCode, Optional[bool]] = {
            DecisionCode.ADOPT: True,
            DecisionCode.REJECT_CANDIDATE: False,
            DecisionCode.INCOMPLETE: None,
            DecisionCode.UNINTERPRETABLE: None,
        }
        if self.candidate_adoption is not expected[self.decision]:
            raise ValueError("candidate_adoption does not match decision")

    def as_dict(self) -> dict[str, Any]:
        return {
            "schema": OUTPUT_SCHEMA,
            "tool": {"name": TOOL_NAME, "version": TOOL_VERSION},
            "candidate": self.facts.candidate,
            "source_revision": self.facts.source_revision,
            "decision": self.decision.value,
            "candidate_adoption": self.candidate_adoption,
            "completeness": self.facts.completeness.as_dict(),
            "facets": {
                name: self.facts.facets()[name].as_dict() for name in _FACET_NAMES
            },
            "reasons": list(self.reasons),
            "provenance": [item.as_dict() for item in self.facts.provenance],
        }


def _facet_reasons(name: str, facet: Facet) -> Tuple[str, ...]:
    return tuple(f"{name}:{code}" for code in facet.reason_codes)


def decide(facts: QualificationFacts) -> DecisionResult:
    """Apply the fixed precedence rules to validated facts.

    Missing evidence wins before any other state.  Custody and readiness
    failures make the evidence uninterpretable; they never reject a candidate.
    Interpretable correctness, performance, and benefit failures reject a
    candidate.  Adoption requires every facet to pass.
    """

    if type(facts) is not QualificationFacts:
        raise TypeError("facts must be QualificationFacts")
    facets = facts.facets()

    missing: list[str] = []
    if facts.completeness.status is CompletenessStatus.MISSING:
        missing.extend(_facet_reasons("completeness", facts.completeness))
    for name, facet in facets.items():
        if facet.status is FacetStatus.MISSING:
            missing.extend(_facet_reasons(name, facet))
    if missing:
        return DecisionResult(
            facts=facts,
            decision=DecisionCode.INCOMPLETE,
            candidate_adoption=None,
            reasons=tuple(missing),
        )

    uninterpretable: list[str] = []
    for name in ("custody", "readiness"):
        facet = facets[name]
        if facet.status in (FacetStatus.FAIL, FacetStatus.UNINTERPRETABLE):
            uninterpretable.extend(_facet_reasons(name, facet))
    for name in ("correctness", "performance", "benefit"):
        facet = facets[name]
        if facet.status is FacetStatus.UNINTERPRETABLE:
            uninterpretable.extend(_facet_reasons(name, facet))
    if uninterpretable:
        return DecisionResult(
            facts=facts,
            decision=DecisionCode.UNINTERPRETABLE,
            candidate_adoption=None,
            reasons=tuple(uninterpretable),
        )

    rejected: list[str] = []
    for name in ("correctness", "performance", "benefit"):
        facet = facets[name]
        if facet.status is FacetStatus.FAIL:
            rejected.extend(_facet_reasons(name, facet))
    if rejected:
        return DecisionResult(
            facts=facts,
            decision=DecisionCode.REJECT_CANDIDATE,
            candidate_adoption=False,
            reasons=tuple(rejected),
        )

    # Missing, uninterpretable, and fail have all been handled above.  The
    # explicit check keeps a future enum extension from silently adopting.
    if not all(facet.status is FacetStatus.PASS for facet in facets.values()):
        raise DecisionInputError("facts contain an unsupported facet state")
    return DecisionResult(
        facts=facts,
        decision=DecisionCode.ADOPT,
        candidate_adoption=True,
        reasons=("all_required_gates_passed",),
    )


def _require_mapping(value: Any, context: str) -> Mapping[str, Any]:
    if type(value) is not dict:
        raise DecisionInputError(f"{context} must be an object")
    if any(type(key) is not str for key in value):
        raise DecisionInputError(f"{context} keys must be strings")
    return value


def _check_keys(value: Mapping[str, Any], expected: Sequence[str], context: str) -> None:
    expected_set = set(expected)
    actual_set = set(value)
    missing = sorted(expected_set - actual_set)
    unknown = sorted(actual_set - expected_set)
    if missing:
        raise DecisionInputError(f"{context} missing keys: {', '.join(missing)}")
    if unknown:
        raise DecisionInputError(f"{context} unknown keys: {', '.join(unknown)}")


def _parse_reason_codes(value: Any, context: str) -> Tuple[str, ...]:
    if type(value) is not list:
        raise DecisionInputError(f"{context}.reason_codes must be an array")
    try:
        return _validate_reason_codes(tuple(value), context)
    except (TypeError, ValueError) as error:
        raise DecisionInputError(str(error)) from error


def _parse_status(value: Any, enum_type: type[Enum], context: str) -> Enum:
    if type(value) is not str:
        raise DecisionInputError(f"{context}.status must be a string")
    try:
        return enum_type(value)
    except ValueError as error:
        allowed = ", ".join(item.value for item in enum_type)
        raise DecisionInputError(
            f"{context}.status must be one of: {allowed}"
        ) from error


def _parse_completeness(value: Any) -> Completeness:
    mapping = _require_mapping(value, "completeness")
    _check_keys(mapping, ("status", "reason_codes"), "completeness")
    status = _parse_status(mapping["status"], CompletenessStatus, "completeness")
    return Completeness(
        status=status,
        reason_codes=_parse_reason_codes(mapping["reason_codes"], "completeness"),
    )


def _parse_facet(value: Any, name: str) -> Facet:
    mapping = _require_mapping(value, name)
    _check_keys(mapping, ("status", "reason_codes"), name)
    status = _parse_status(mapping["status"], FacetStatus, name)
    return Facet(status=status, reason_codes=_parse_reason_codes(mapping["reason_codes"], name))


def _parse_provenance(value: Any) -> Tuple[ProvenanceRef, ...]:
    if type(value) is not list or not value:
        raise DecisionInputError("provenance must be a nonempty array")
    result = []
    for index, item in enumerate(value):
        context = f"provenance[{index}]"
        mapping = _require_mapping(item, context)
        _check_keys(mapping, ("kind", "path", "sha256"), context)
        try:
            result.append(
                ProvenanceRef(
                    kind=mapping["kind"],
                    path=mapping["path"],
                    sha256=mapping["sha256"],
                )
            )
        except (TypeError, ValueError) as error:
            raise DecisionInputError(str(error)) from error
    return tuple(result)


def facts_from_mapping(value: Any) -> QualificationFacts:
    """Parse a strict input mapping; input mappings contain no verdict field."""

    mapping = _require_mapping(value, "input")
    _check_keys(
        mapping,
        ("schema", "candidate", "source_revision", "completeness", "facets", "provenance"),
        "input",
    )
    if mapping["schema"] != INPUT_SCHEMA:
        raise DecisionInputError(f"input.schema must equal {INPUT_SCHEMA}")
    for name in ("candidate", "source_revision"):
        if type(mapping[name]) is not str or not mapping[name]:
            raise DecisionInputError(f"input.{name} must be a nonempty string")
    facets_mapping = _require_mapping(mapping["facets"], "facets")
    _check_keys(facets_mapping, _FACET_NAMES, "facets")
    try:
        return QualificationFacts(
            candidate=mapping["candidate"],
            source_revision=mapping["source_revision"],
            completeness=_parse_completeness(mapping["completeness"]),
            custody=_parse_facet(facets_mapping["custody"], "facets.custody"),
            correctness=_parse_facet(
                facets_mapping["correctness"], "facets.correctness"
            ),
            readiness=_parse_facet(facets_mapping["readiness"], "facets.readiness"),
            performance=_parse_facet(
                facets_mapping["performance"], "facets.performance"
            ),
            benefit=_parse_facet(facets_mapping["benefit"], "facets.benefit"),
            provenance=_parse_provenance(mapping["provenance"]),
        )
    except (TypeError, ValueError) as error:
        raise DecisionInputError(str(error)) from error


def _reject_duplicate_keys(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for key, value in pairs:
        if key in result:
            raise DecisionInputError(f"duplicate JSON key: {key}")
        result[key] = value
    return result


def _reject_nonfinite(value: str) -> None:
    raise DecisionInputError(f"non-finite JSON number is not allowed: {value}")


def load_facts(path: Path) -> QualificationFacts:
    """Load and validate one facts document without producing any output."""

    try:
        text = path.read_text(encoding="utf-8")
        value = json.loads(
            text,
            object_pairs_hook=_reject_duplicate_keys,
            parse_constant=_reject_nonfinite,
        )
    except DecisionInputError:
        raise
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise DecisionInputError(f"cannot read input {path}: {error}") from error
    return facts_from_mapping(value)


def _json_bytes(value: Mapping[str, Any]) -> bytes:
    try:
        return (
            json.dumps(value, indent=2, sort_keys=True, allow_nan=False)
            + "\n"
        ).encode("utf-8")
    except (TypeError, ValueError) as error:
        raise DecisionOutputError(f"cannot encode decision JSON: {error}") from error


def write_result(path: Path, result: DecisionResult) -> None:
    """Write a result exactly once, refusing an existing file or symlink."""

    payload = _json_bytes(result.as_dict())
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        flags = os.O_WRONLY | os.O_CREAT | os.O_EXCL
        if hasattr(os, "O_NOFOLLOW"):
            flags |= os.O_NOFOLLOW
        if hasattr(os, "O_CLOEXEC"):
            flags |= os.O_CLOEXEC
        descriptor = os.open(os.fspath(path), flags, 0o644)
    except OSError as error:
        raise DecisionOutputError(f"refusing output {path}: {error}") from error

    try:
        with os.fdopen(descriptor, "wb") as handle:
            descriptor = -1
            handle.write(payload)
            handle.flush()
            os.fsync(handle.fileno())
    except OSError as error:
        if descriptor >= 0:
            os.close(descriptor)
        raise DecisionOutputError(f"cannot write output {path}: {error}") from error


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input", required=True, type=Path, help="facts JSON input")
    parser.add_argument(
        "--json-out", required=True, type=Path, help="fresh decision JSON output"
    )
    return parser


def main(argv: Optional[Sequence[str]] = None) -> int:
    args = build_parser().parse_args(argv)
    try:
        facts = load_facts(args.input)
        result = decide(facts)
        write_result(args.json_out, result)
    except (DecisionInputError, DecisionOutputError) as error:
        print(f"decision rejected: {error}", file=sys.stderr)
        return 2
    print(
        json.dumps(
            {
                "decision": result.decision.value,
                "candidate_adoption": result.candidate_adoption,
            },
            sort_keys=True,
        )
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
