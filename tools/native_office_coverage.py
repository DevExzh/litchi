#!/usr/bin/env python3
"""Report native-fixture test scopes separately from historical authoring evidence.

The index points to real Rust tests. Presence in the index means only that the
listed assertions exist, not that an entire feature row is complete or that a
test passed in this checkout. Run the referenced Cargo targets to establish that.
Unmapped fixtures and rows remain visible in every generated report.
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


def report() -> dict:
    manifest = json.loads((CORPUS / "manifest.json").read_text())
    inventory = json.loads((CORPUS / "coverage.json").read_text())
    review = json.loads((CORPUS / "final-coverage-review.json").read_text())
    index = json.loads(INDEX.read_text())
    fixtures = {item["path"]: [] for item in manifest["files"]}
    rows = {item["id"]: [] for item in inventory["feature_inventory"]}
    dispositions = {item["id"]: item for item in review["pending_rows"]}
    test_ids = set()
    for test in index["tests"]:
        identity = f"{test['crate']}::{test['target']}::{test['name']}"
        if identity in test_ids:
            raise ValueError(f"duplicate test: {identity}")
        test_ids.add(identity)
        for component in (test["crate"], test["target"], test["name"]):
            if not re.fullmatch(r"[a-z][a-z0-9_-]*", component):
                raise ValueError(f"invalid Rust test identity: {identity}")
        source = ROOT / "crates" / test["crate"] / "tests" / f"{test['target']}.rs"
        text = source.read_text()
        if not re.search(r"#\[test\]\s*fn\s+" + re.escape(test["name"]) + r"\s*\(", text):
            raise ValueError(f"missing test: {identity} in {source}")
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
        row = {
            "id": item["id"],
            "format": item["format"],
            "feature": item["feature"],
            "native_fixture_status": item["fixture_status"],
            "automated_coverage": "partial" if rows[item["id"]] else "uncovered",
            "tests": rows[item["id"]],
        }
        if item["id"] in dispositions:
            disposition = dispositions[item["id"]]
            row["native_pending_classification"] = disposition["classification"]
            row["native_pending_rationale"] = disposition["rationale"]
        feature_rows.append(row)
    return {
        "schema_version": 1,
        "scope": "Test assertions indexed, not whole-feature completion or an execution receipt.",
        "summary": {
            "fixtures": len(fixtures),
            "fixtures_with_tests": sum(bool(tests) for tests in fixtures.values()),
            "feature_rows": len(rows),
            "feature_rows_with_partial_tests": sum(bool(tests) for tests in rows.values()),
            "indexed_tests": len(test_ids),
            "native_fixture_statuses": dict(Counter(item["fixture_status"] for item in inventory["feature_inventory"])),
        },
        "fixtures": [{"path": path, "tests": tests} for path, tests in fixtures.items()],
        "feature_rows": feature_rows,
        "tests": index["tests"],
    }


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--summary", action="store_true", help="print counts instead of the complete inventory")
    args = parser.parse_args()
    result = report()
    print(json.dumps(result["summary"] if args.summary else result, indent=2))


if __name__ == "__main__":
    main()
