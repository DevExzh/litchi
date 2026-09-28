#!/usr/bin/env python3
"""Record 0803 diagnostics without granting workflow advancement.

The 0803 packet compares the rejected 0802 candidate after-state with one
single-factor relocation of its exact-empty check.  It is deliberately outside
the qualification workflow: this reader records measured diagnostic changes,
but the advancement and adoption fields are permanently false.
"""

from __future__ import annotations

import argparse
from typing import Any

import custody as c


p = c.P
SCHEMA = "litchi.performance.0803.diagnostic-decision.v1"
PLAN_SCHEMA = "litchi.performance.0803.v1"


def diagnostic_changes(analysis: dict[str, Any]) -> dict[str, Any]:
    """Retain paired native and Callgrind changes as diagnostics only."""

    native = analysis.get("native", {}).get("analysis", {})
    paired = native.get("paired_by_case_mode", {})
    rows: list[dict[str, Any]] = []
    if isinstance(paired, dict):
        for key in sorted(paired):
            row = paired[key]
            if not isinstance(row, dict):
                continue
            rows.append({
                "case": row.get("case"),
                "mode": row.get("mode"),
                "ratio_median": row.get("ratio_median"),
                "change_percent_median": row.get("change_percent_median"),
                "bootstrap": row.get("bootstrap"),
                "diagnostic_regression": row.get("diagnostic_regression"),
            })

    profiles = analysis.get("profiles", {}).get("analysis", {})
    counter_summary = profiles.get("counter_summary", {})
    paired_profiles = counter_summary.get("paired_by_case_mode", {})
    profile_rows: list[dict[str, Any]] = []
    if isinstance(paired_profiles, dict):
        for key in sorted(paired_profiles):
            row = paired_profiles[key]
            if not isinstance(row, dict):
                continue
            profile_rows.append({"key": key, **row})
    return {
        "native": rows,
        "profile": profile_rows,
    }


def build_result() -> dict[str, Any]:
    plan = c.read(p / "plan.json")
    assert plan["schema"] == PLAN_SCHEMA
    assert plan["analysis"]["bootstrap_seed"] == 803080
    assert "preflight_decision" not in plan
    policy = plan.get("diagnostic_decision")
    assert isinstance(policy, dict)
    assert policy.get("advance_to_workflow_trials") is False
    assert policy.get("production_adoption") is False
    assert policy.get("no_production_baseline_comparison") is True

    analysis = c.read(p / "analysis.json")
    audit = c.read(p / "root-native-audit.json")
    changes = diagnostic_changes(analysis)
    assert len(changes["native"]) == 78
    assert len(changes["profile"]) == 78
    return {
        "schema": SCHEMA,
        "comparison": policy.get("comparison"),
        "single_factor": policy.get("single_factor"),
        "no_production_baseline_comparison": True,
        "advance_to_workflow_trials": False,
        "production_adoption": False,
        "diagnostic_changes": changes,
        "analysis": c.artifact(p / "analysis.json"),
        "independent_audit": c.artifact(p / "root-native-audit.json"),
        "reports": audit.get("reports"),
        "samples": audit.get("samples"),
    }


def main() -> int:
    parser = argparse.ArgumentParser()
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument("--write", action="store_true")
    group.add_argument("--check", action="store_true")
    args = parser.parse_args()
    result = build_result()
    output = p / "decision.json"
    if args.write:
        assert not output.exists(), output
        c.write(output, result)
    else:
        assert c.read(output) == result
    print("0803 diagnostic decision recorded; workflow advancement and production adoption remain false")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
