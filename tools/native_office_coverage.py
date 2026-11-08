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
MATRIX_RECONCILIATION = ROOT / "tools/native_office_matrix_reconciliation.json"


def matrix_rows(text: str, source: str) -> dict:
    """Read checklist labels/statuses without putting test metadata in matrices."""
    rows = {}
    columns = None
    separator = False
    for number, line in enumerate(text.splitlines(), 1):
        if not line.strip().startswith("|"):
            if separator:
                raise ValueError(f"missing matrix header separator: {source}:{number}")
            columns = None
            continue
        cells = [cell.strip() for cell in line.strip().strip("|").split("|")]
        location = f"{source}:{number}"
        if columns is None:
            if not cells[0].startswith("Feature") or cells[1:] not in (
                ["Status", "Read", "Write"], ["Read", "Write"]
            ):
                raise ValueError(f"invalid matrix header: {location}")
            columns = [cell.lower() for cell in cells[1:]]
            separator = True
        elif separator:
            if len(cells) != len(columns) + 1 or not all(re.fullmatch(r":?-+:?", cell) for cell in cells):
                raise ValueError(f"invalid matrix header separator: {location}")
            separator = False
        else:
            if len(cells) != len(columns) + 1 or not cells[0] or any(
                cell not in {"✅", "🟡", "❌", "N/A"} for cell in cells[1:]
            ):
                raise ValueError(f"invalid matrix row: {location}")
            if cells[0] in rows:
                raise ValueError(f"duplicate matrix feature: {location}: {cells[0]}")
            rows[cells[0]] = dict(zip(columns, cells[1:]))
    if separator or not rows:
        raise ValueError(f"incomplete or empty matrix: {source}")
    return rows


def reconcile_matrices(inventory: list[dict], document: dict) -> dict:
    """Require explicit accounting for drift from the immutable authoring inventory."""
    if document.get("schema_version") != 1:
        raise ValueError("unsupported native matrix reconciliation schema")
    historical = {row["id"]: row for row in inventory}
    sources = {row["source_matrix"]: row["format"] for row in inventory}
    current = {source: matrix_rows((ROOT / source).read_text(), source) for source in sources}
    for source, features in current.items():
        columns = {"read", "write"} if sources[source] == "rtf" else {"status", "read", "write"}
        if any(set(status) != columns for status in features.values()):
            raise ValueError(f"changed matrix status columns: {source}")

    def declarations(key: str) -> dict:
        entries = {}
        for entry in document[key]:
            row = entry["feature_row"]
            if row not in historical or row in entries:
                raise ValueError(f"unknown or duplicate {key} row: {row}")
            if not isinstance(entry.get("reason"), str) or not entry["reason"].strip():
                raise ValueError(f"missing {key} reason: {row}")
            entries[row] = entry
        return entries

    renamed = declarations("renamed_rows")
    consolidated = declarations("consolidated_rows")
    changes = declarations("status_changes")
    if set(consolidated) & (set(renamed) | set(changes)):
        raise ValueError("consolidated rows cannot also be renamed or change status")
    bindings = {}
    rows = []

    def bind(identity: str, format: str, source: str, feature: str, status: dict, historical_id: str | None) -> None:
        key = (source, feature)
        if source not in current or feature not in current[source]:
            raise ValueError(f"missing current matrix feature: {identity}: {feature}")
        if key in bindings:
            raise ValueError(f"duplicate current matrix binding: {identity}: {feature}")
        if status != current[source][feature]:
            raise ValueError(f"unreconciled matrix status: {identity}: {feature}")
        bindings[key] = identity
        rows.append({
            "id": identity, "format": format, "source_matrix": source,
            "feature": feature, "status": status, "historical_feature_row": historical_id,
        })

    for row in inventory:
        identity, source, feature = row["id"], row["source_matrix"], row["feature"]
        if identity in consolidated:
            target = consolidated[identity]["into_feature_row"]
            if (target not in historical or target in consolidated or
                    historical[target]["source_matrix"] != source or feature in current[source]):
                raise ValueError(f"invalid matrix consolidation: {identity}")
            continue
        if identity in renamed:
            feature = renamed[identity]["current_feature"]
            if feature == row["feature"]:
                raise ValueError(f"redundant matrix rename: {identity}")
        if feature not in current[source]:
            raise ValueError(f"missing current matrix feature: {identity}: {feature}")
        historical_status = {
            # The immutable inventory includes one old prose-qualified status.
            column: row["matrix_status" if column == "status" else column].split()[0]
            for column in current[source][feature]
        }
        expected = historical_status
        if identity in changes:
            expected = changes[identity]["current_status"]
            if expected == historical_status:
                raise ValueError(f"redundant matrix status change: {identity}")
        bind(identity, row["format"], source, feature, expected, identity)

    additional_ids = set()
    for entry in document["additional_rows"]:
        identity, source = entry["id"], entry["source_matrix"]
        if (not re.fullmatch(r"[a-z]+-current-[a-z0-9-]+", identity) or
                identity in historical or identity in additional_ids):
            raise ValueError(f"invalid or duplicate additional matrix id: {identity}")
        if source not in sources or entry["format"] != sources[source]:
            raise ValueError(f"invalid additional matrix source: {identity}")
        if not isinstance(entry.get("scope_limit"), str) or not entry["scope_limit"].strip():
            raise ValueError(f"missing additional matrix scope: {identity}")
        additional_ids.add(identity)
        bind(identity, entry["format"], source, entry["feature"], entry["current_status"], None)
        rows[-1]["scope_limit"] = entry["scope_limit"]

    missing = {(source, feature) for source, features in current.items() for feature in features} - bindings.keys()
    if missing:
        raise ValueError(f"unreconciled current matrix features: {sorted(missing)}")
    return {
        "rows": rows,
        "renamed_rows": list(renamed.values()),
        "consolidated_rows": list(consolidated.values()),
        "status_changes": list(changes.values()),
    }


def test_identity(test: dict) -> str:
    """Validate a Rust check reference without treating its existence as a pass."""
    identity = f"{test['crate']}::{test['target']}::{test['name']}"
    for component in (test["crate"], test["target"], test["name"]):
        if not re.fullmatch(r"[a-z][a-z0-9_-]*", component):
            raise ValueError(f"invalid Rust test identity: {identity}")
    source = ROOT / "crates" / test["crate"] / "tests" / f"{test['target']}.rs"
    if not re.search(
        r"(?:#\[test\]\s*fn|#\[tokio::test\]\s*async\s+fn)\s+"
        + re.escape(test["name"]) + r"\s*\(",
        source.read_text(),
    ):
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
    if index.get("schema_version") != 1:
        raise ValueError("unsupported native test index schema")
    current_matrix = reconcile_matrices(
        inventory["feature_inventory"], json.loads(MATRIX_RECONCILIATION.read_text())
    )
    additional = {row["id"]: [] for row in current_matrix["rows"] if row["historical_feature_row"] is None}
    fixtures = {item["path"]: [] for item in manifest["files"]}
    rows = {item["id"]: [] for item in inventory["feature_inventory"]}
    reviewed = evidence_reviews(json.loads(EVIDENCE_REVIEWS.read_text()), rows, fixtures)
    fixture_checks = {path: set() for path in fixtures}
    for evidence in reviewed.values():
        checks = {f"{check['crate']}::{check['target']}::{check['name']}" for check in evidence["checks"]}
        for fixture in evidence["fixtures"]:
            fixture_checks[fixture].update(checks)
    dispositions = {item["id"]: item for item in review["pending_rows"]}
    test_ids = set()
    for test in index["tests"]:
        identity = test_identity(test)
        if identity in test_ids:
            raise ValueError(f"duplicate test: {identity}")
        test_ids.add(identity)
        if not test["assertions"] or not test["scope_limit"]:
            raise ValueError(f"assertions and scope limit are required: {identity}")
        if not test["fixtures"] or not (test["feature_rows"] or test.get("current_features")):
            raise ValueError(f"fixture and feature bindings are required: {identity}")
        for fixture in test["fixtures"]:
            if fixture not in fixtures:
                raise ValueError(f"unknown fixture {fixture}: {identity}")
            fixtures[fixture].append(identity)
        for row in test["feature_rows"]:
            if row not in rows:
                raise ValueError(f"unknown feature row {row}: {identity}")
            rows[row].append(identity)
        for row in test.get("current_features", []):
            if row not in additional:
                raise ValueError(f"unknown additional matrix feature {row}: {identity}")
            additional[row].append(identity)

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
    by_id = {row["id"]: row for row in feature_rows}
    for row in current_matrix["rows"]:
        historical_id = row["historical_feature_row"]
        if historical_id is not None:
            historical_row = by_id[historical_id]
            for field in ("tests", "automated_coverage", "verification_gap"):
                row[field] = historical_row[field]
        else:
            row["tests"] = additional[row["id"]]
            row["automated_coverage"] = "partial" if row["tests"] else "uncovered"
            row["verification_gap"] = (
                "partial_assertions_require_scope_review" if row["tests"]
                else "current_feature_needs_semantic_tests"
            )
        if row["tests"] and set(row["status"].values()) <= {"❌", "N/A"}:
            raise ValueError(f"unsupported matrix feature cannot claim semantic assertions: {row['id']}")
    return {
        "schema_version": 1,
        "scope": "Semantic assertions, reviewed fixture gaps and explicit scope boundaries are separate from historical authoring evidence; none implies whole-feature completion or a test execution receipt.",
        "summary": {
            "fixtures": len(fixtures),
            "fixtures_with_tests": sum(bool(tests) for tests in fixtures.values()),
            "fixtures_with_evidence_checks": sum(bool(checks) for checks in fixture_checks.values()),
            "fixtures_with_any_checks": sum(bool(tests or fixture_checks[path]) for path, tests in fixtures.items()),
            "feature_rows": len(rows),
            "feature_rows_with_partial_tests": sum(bool(tests) for tests in rows.values()),
            "indexed_tests": len(test_ids),
            "current_matrix_rows": len(current_matrix["rows"]),
            "current_matrix_rows_with_partial_tests": sum(bool(row["tests"]) for row in current_matrix["rows"]),
            "current_only_rows": len(additional),
            "renamed_matrix_rows": len(current_matrix["renamed_rows"]),
            "consolidated_historical_rows": len(current_matrix["consolidated_rows"]),
            "changed_matrix_statuses": len(current_matrix["status_changes"]),
            "reviewed_native_fixture_gaps": sum(review["classification"] != "explicit_scope_boundary" for review in reviewed.values()),
            "reviewed_explicit_scope_boundaries": sum(review["classification"] == "explicit_scope_boundary" for review in reviewed.values()),
            "native_fixture_statuses": dict(Counter(item["fixture_status"] for item in inventory["feature_inventory"])),
            "verification_gaps": dict(Counter(row["verification_gap"] for row in feature_rows)),
        },
        "fixtures": [
            {"path": path, "tests": tests, "evidence_checks": sorted(fixture_checks[path])}
            for path, tests in fixtures.items()
        ],
        "feature_rows": feature_rows,
        "current_matrix": current_matrix,
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
        output.extend(row for row in result["current_matrix"]["rows"] if row["historical_feature_row"] is None and not row["tests"])
    print(json.dumps(output, indent=2))


if __name__ == "__main__":
    main()
