#!/usr/bin/env python3
"""Synthetic controls for the 0964 selector; never launches a workload."""
from __future__ import annotations

import copy
import importlib.util
import json
from pathlib import Path

HERE = Path(__file__).resolve().parent
SPEC = importlib.util.spec_from_file_location("selector_0964", HERE / "select_screen_0964_v1.py")
assert SPEC and SPEC.loader
MODULE = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(MODULE)


def fixture(value: float = 1.0) -> dict:
    def native_metrics() -> dict:
        keys = MODULE.NATIVE_REQUIRED
        return {key: {"pair_count": 6, "undefined_count": 0, "median": value} for key in keys}

    def allocator_metrics() -> dict:
        keys = MODULE.ALLOCATOR_REQUIRED
        return {key: {"pair_count": 3, "undefined_count": 0, "median": value} for key in keys}

    return {
        "status": "screened",
        "native": {"cases": {case: {"metrics": native_metrics()} for case in MODULE.CASES}},
        "allocator": {"cases": {case: {"metrics": allocator_metrics()} for case in MODULE.CASES}},
    }


def expect(name: str, report: dict, decision: str, needle: str | None = None) -> dict:
    result = MODULE.select(report)
    assert result["decision"] == decision, (name, result)
    if needle:
        assert any(needle in item for item in result["stops"]), (name, result)
    return {"name": name, "status": "pass", "decision": result["decision"]}


def main() -> int:
    checks = [
        expect("signal-equality-0.95", fixture(0.95), "require_full_confirmation_and_regression_review"),
        expect("safety-equality-1.05", fixture(1.05), "reject_screen", "No qualifying metric"),
        expect("no-signal", fixture(1.0), "reject_screen", "No qualifying metric"),
        expect("safety-stop", fixture(1.050001), "reject_screen", "safety median"),
    ]
    invalid = fixture(0.95)
    invalid["native"]["cases"][MODULE.CASES[0]]["metrics"]["total_ns.p99"] = {
        "pair_count": 5,
        "undefined_count": 0,
        "median": 0.95,
    }
    checks.append(expect("invalid-pair-count", invalid, "reject_screen", "invalid or undefined"))
    requested_only = fixture(1.0)
    for case in MODULE.CASES:
        requested_only["allocator"]["cases"][case]["metrics"]["allocated_bytes"]["median"] = 0.90
        requested_only["allocator"]["cases"][case]["metrics"]["allocation_calls"]["median"] = 0.90
    checks.append(expect("counter-only-no-signal", requested_only, "reject_screen", "No qualifying metric"))
    print(json.dumps({"schema": "litchi.performance.0964.selector-controls.v1", "status": "pass", "checks": checks}, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
