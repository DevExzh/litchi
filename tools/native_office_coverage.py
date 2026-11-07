#!/usr/bin/env python3
"""Report native-fixture test scopes separately from historical authoring evidence.

The index points to real Rust tests. Presence in the index means only that the
listed assertions exist, not that an entire feature row is complete or that a
test passed in this checkout. Run the referenced Cargo targets to establish that.
Unmapped fixtures and rows remain visible in every generated report.
Evidence-only checks live in a separate review index. They can correct a
historical fixture claim or document an explicit scope boundary, but never
count as semantic feature coverage.
"""

from __future__ import annotations

import argparse
from collections import Counter
import json
from pathlib import Path
import re


ROOT = Path(__file__).resolve().parents[1]
CORPUS = ROOT / "test-data/microsoft-office"
INDEX = ROOT / "tools/native_office_tests.json"
EVIDENCE_REVIEWS = ROOT / "tools/native_office_evidence_reviews.json"


def test_identity(test: dict) -> str:
    """Validate a Rust check reference without treating its existence as a pass."""
    identity = f"{test['crate']}::{test['target']}::{test['name']}"
    for component in (test["crate"], test["target"], test["name"]):
        if not re.fullmatch(r"[a-z][a-z0-9_-]*", component):
            raise ValueError(f"invalid Rust test identity: {identity}")
    source = ROOT / "crates" / test["crate"] / "tests" / f"{test['target']}.rs"
    if not re.search(r"#\[test\]\s*fn\s+" + re.escape(test["name"]) + r"\s*\(", source.read_text()):
        raise ValueError(f"missing test: {identity} in {source}")
    return identity


def evidence_reviews(document: dict, rows: dict, fixtures: dict) -> dict:
    """Validate evidence corrections separately from positive semantic assertions."""
    if document.get("schema_version") != 1:
        raise ValueError("unsupported native evidence review schema")
    reviewed = {}
    for review in document["reviews"]:
        row = review["feature_row"]
        if row not in rows or row in reviewed:
            raise ValueError(f"unknown or duplicate evidence review row: {row}")
        if review["classification"] not in {
            "fixture_feature_mismatch", "native_derivative_outside_semantic_profile",
            "explicit_scope_boundary"
        }:
            raise ValueError(f"unknown evidence review classification: {row}")
        fields = ["reason", "required_evidence", "observed_evidence", "scope_limit"]
        if review["classification"] == "explicit_scope_boundary":
            fields.extend(("supported_scope", "unsupported_scope"))
        for field in fields:
            if not isinstance(review.get(field), str) or not review[field].strip():
                raise ValueError(f"missing evidence review {field}: {row}")
        if not review["fixtures"] or not review["checks"] or not review["evidence_refs"]:
            raise ValueError(f"evidence review requires fixtures, checks and sources: {row}")
        for fixture in review["fixtures"]:
            if fixture not in fixtures:
                raise ValueError(f"unknown evidence review fixture: {fixture}")
        for reference in review["evidence_refs"]:
            if not (ROOT / reference).is_file():
                raise ValueError(f"missing evidence review source: {reference}")
        for check in review["checks"]:
            test_identity(check)
        reviewed[row] = review
    return reviewed


def verification_gap(item: dict, tests: list[str], disposition: dict | None, evidence: dict | None = None) -> str:
    """Classify the next review, never infer completion from fixture presence."""
    if tests:
        return "partial_assertions_require_scope_review"
    if evidence is not None:
        if evidence["classification"] == "explicit_scope_boundary":
            return "reviewed_explicit_scope_boundary"
        return "reviewed_native_fixture_gap_requires_followup"
    status = item["fixture_status"]
    if status == "verified_native_fixture":
        return "native_evidence_needs_semantic_tests"
    if status == "library_boundary":
        return "unsupported_boundary_needs_review"
    if status == "not_applicable":
        return "not_applicable"
    if status == "pending" and disposition is not None:
        return {
            "operation_only": "operation_needs_verification",
            "native_bounded_evidence_operation_unverified": "bounded_native_evidence_needs_verification",
            "observed_native_ui_limitation": "native_ui_limitation_needs_review",
            "external_handoff": "external_fixture_handoff_needs_review",
        }[disposition["classification"]]
    raise ValueError(f"unclassified native row: {item['id']}")


def report() -> dict:
    manifest = json.loads((CORPUS / "manifest.json").read_text())
    inventory = json.loads((CORPUS / "coverage.json").read_text())
    review = json.loads((CORPUS / "final-coverage-review.json").read_text())
    index = json.loads(INDEX.read_text())
    fixtures = {item["path"]: [] for item in manifest["files"]}
    rows = {item["id"]: [] for item in inventory["feature_inventory"]}
    reviewed = evidence_reviews(json.loads(EVIDENCE_REVIEWS.read_text()), rows, fixtures)
    dispositions = {item["id"]: item for item in review["pending_rows"]}
    test_ids = set()
    for test in index["tests"]:
        identity = test_identity(test)
        if identity in test_ids:
            raise ValueError(f"duplicate test: {identity}")
        test_ids.add(identity)
        if not test["assertions"] or not test["scope_limit"]:
            raise ValueError(f"assertions and scope limit are required: {identity}")
        if not test["fixtures"] or not test["feature_rows"]:
            raise ValueError(f"fixture and feature bindings are required: {identity}")
        for fixture in test["fixtures"]:
            if fixture not in fixtures:
                raise ValueError(f"unknown fixture {fixture}: {identity}")
            fixtures[fixture].append(identity)
        for row in test["feature_rows"]:
            if row not in rows:
                raise ValueError(f"unknown feature row {row}: {identity}")
            rows[row].append(identity)

    feature_rows = []
    for item in inventory["feature_inventory"]:
        evidence = reviewed.get(item["id"])
        if evidence is not None and rows[item["id"]]:
            raise ValueError(f"evidence review conflicts with semantic assertions: {item['id']}; reconcile the review first")
        row = {
            "id": item["id"],
            "format": item["format"],
            "feature": item["feature"],
            "native_fixture_status": item["fixture_status"],
            "automated_coverage": "partial" if rows[item["id"]] else "uncovered",
            "tests": rows[item["id"]],
            "verification_gap": verification_gap(item, rows[item["id"]], dispositions.get(item["id"]), evidence),
            "historical_support_note": item.get("support_note", ""),
            "source_matrix": item["source_matrix"],
        }
        if evidence is not None:
            row["evidence_review"] = evidence
        if item["id"] in dispositions:
            disposition = dispositions[item["id"]]
            row["native_pending_classification"] = disposition["classification"]
            row["native_pending_rationale"] = disposition["rationale"]
            row["historical_evidence_refs"] = disposition["evidence_refs"]
        feature_rows.append(row)
    return {
        "schema_version": 1,
        "scope": "Semantic assertions, reviewed fixture gaps and explicit scope boundaries are separate from historical authoring evidence; none implies whole-feature completion or a test execution receipt.",
        "summary": {
            "fixtures": len(fixtures),
            "fixtures_with_tests": sum(bool(tests) for tests in fixtures.values()),
            "feature_rows": len(rows),
            "feature_rows_with_partial_tests": sum(bool(tests) for tests in rows.values()),
            "indexed_tests": len(test_ids),
            "reviewed_native_fixture_gaps": sum(review["classification"] != "explicit_scope_boundary" for review in reviewed.values()),
            "reviewed_explicit_scope_boundaries": sum(review["classification"] == "explicit_scope_boundary" for review in reviewed.values()),
            "native_fixture_statuses": dict(Counter(item["fixture_status"] for item in inventory["feature_inventory"])),
            "verification_gaps": dict(Counter(row["verification_gap"] for row in feature_rows)),
        },
        "fixtures": [{"path": path, "tests": tests} for path, tests in fixtures.items()],
        "feature_rows": feature_rows,
        "tests": index["tests"],
    }


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    mode = parser.add_mutually_exclusive_group()
    mode.add_argument("--summary", action="store_true", help="print counts instead of the complete inventory")
    mode.add_argument("--gaps", action="store_true", help="print unasserted feature rows and their historical evidence; these are review tasks, not accepted exclusions")
    args = parser.parse_args()
    result = report()
    output = result["summary"] if args.summary else result
    if args.gaps:
        output = [row for row in result["feature_rows"] if not row["tests"]]
    print(json.dumps(output, indent=2))


if __name__ == "__main__":
    main()
