#!/usr/bin/env python3
"""Synthetic boundary checks for the inert 0961 root selector."""
from __future__ import annotations

import importlib.util
import sys
from pathlib import Path

sys.dont_write_bytecode = True
PATH = Path(__file__).with_name("select_screen_0961_v1.py")
spec = importlib.util.spec_from_file_location("selector_0961", PATH)
assert spec and spec.loader
selector = importlib.util.module_from_spec(spec)
spec.loader.exec_module(selector)


def metric(value: float, pairs: int, undefined: int = 0) -> dict[str, object]:
    return {"median": value, "pair_count": pairs, "undefined_count": undefined}


def report(native_value: float = 1.0, rss: float = 1.0, alloc_peak: float = 1.0,
           alloc_rss: float = 1.0, *, undefined: bool = False,
           requested_signal: bool = False) -> dict[str, object]:
    native_cases = {}
    alloc_cases = {}
    for case in selector.CASES:
        native_cases[case] = {"metrics": {
            key: metric(native_value, 6, 1 if undefined and key == "total_ns.p50" else 0)
            for key in (*selector.PRIMARY, "max_rss_kib")}}
        native_cases[case]["metrics"]["max_rss_kib"] = metric(rss, 6)
        alloc_cases[case] = {"metrics": {
            "allocated_bytes": metric(0.5 if requested_signal else 1.0, 3),
            "allocation_calls": metric(0.5 if requested_signal else 1.0, 3),
            "region_peak_live_bytes": metric(alloc_peak, 3),
            "core_combined_with_drain.allocated_bytes": metric(1.0, 3),
            "max_rss_kib": metric(alloc_rss, 3),
        }}
    return {"status": "screened", "native": {"cases": native_cases}, "allocator": {"cases": alloc_cases}}


def must(name: str, value: dict[str, object], decision: str, needle: str | None = None) -> None:
    got = selector.select(value)
    assert got["decision"] == decision, (name, got)
    if needle is not None:
        assert any(needle in item for item in got["stops"]), (name, got["stops"])


def main() -> int:
    # Both 0.95 signal and 1.05 safety boundaries are inclusive/defined.
    must("equality-boundaries", report(native_value=0.95, rss=1.05, alloc_peak=1.05, alloc_rss=1.05),
         "require_full_confirmation_and_regression_review")
    must("qualifying-native", report(native_value=0.95), "require_full_confirmation_and_regression_review")
    must("invalid-metric", report(undefined=True), "reject_screen", "invalid or undefined")
    must("native-rss-stop", report(native_value=1.0, rss=1.051), "reject_screen", "native/medium-one-edit/max_rss_kib")
    must("allocator-rss-stop", report(alloc_rss=1.051), "reject_screen", "allocator/medium-one-edit/max_rss_kib")
    must("requested-bytes-not-signal", report(requested_signal=True), "reject_screen", "No qualifying metric")
    print("check_selector_0961_v1: pass (6 synthetic controls)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
