#!/usr/bin/env python3
"""Freeze the reviewed, fixed-population 0885 protocol before any capture."""
import datetime
import json
import os
from pathlib import Path

import capture

PACKET = Path(__file__).resolve().parent

def read(name):
    return json.loads((PACKET / name).read_text())

def main():
    build = read("build.json")
    review = read("methodology-review.json")
    checks = read("helper-checks.json")
    if build["status"] != "pass" or review["status"] != "pass" or checks["exit_code"] != 0:
        raise ValueError("build, helper checks, and methodology review must pass before freezing")
    if 12 not in os.sched_getaffinity(0):
        raise ValueError("CPU 12 is unavailable")
    files = {
        "build": "build.json", "source_manifest": "source.json",
        "source_archive": "source.tar.gz", "capture_helper": "capture.py",
        "verify_helper": "verify.py", "test_helper": "test_helpers.py",
        "review": "methodology-review.json", "environment": "environment.json",
        "taskset": "/usr/bin/taskset", "helper_checks": "helper-checks.json",
        "build_helper": "build.py", "freeze_helper": "freeze.py",
    }
    plan = {
        "schema": capture.SCHEMA,
        "frozen_utc": datetime.datetime.now(datetime.timezone.utc).isoformat(),
        "case": capture.CASE, "profile": "release", "cpu_affinity": "12",
        "cache": "warm", "fresh_child": True, "fail_fast": True,
        "schedule": ["native", "allocator"], "build_json": "build.json",
        "revision": build["revision"],
        "archive": {
            "sha256": capture.ARCHIVE_SHA256, "bytes": capture.ARCHIVE_BYTES,
            "oracle_sha256": capture.ORACLE_SHA256,
            "marker_refusal_census_sha256": capture.CENSUS_SHA256,
            "entry_count": 4096, "numeric_count": 3686, "text_count": 410,
        },
        "lanes": {
            "native": {"name": "native", "mode": "native", "children": 12, "warmup": 3, "samples": 30},
            "allocator": {"name": "allocator", "mode": "allocator", "children": 6, "warmup": 3, "samples": 15},
        },
        "binaries": build["binaries"],
        "bindings": {key: capture.descriptor(PACKET / name) for key, name in files.items()},
        "analysis": {
            "comparison": None, "pool_children": False,
            "native_p50_spread_ratio_flag_above": 1.20,
            "slow_rows": "retain_all_without_retry_or_replacement",
            "statistics": "per-child floored-midpoint p50; nearest-rank p95/p99; across-child median and range",
            "allocator_elapsed": "separate_instrumented_observation",
            "peak_rss": "per-child maximum process-lifetime high-water vector; across-child median and range",
            "throughput": "4096 cells divided by each native child p50",
        },
        "host_scope": "CPU affinity only; no exclusive-host or physical-cold guarantee",
        "load_average_at_freeze": list(os.getloadavg()),
    }
    capture.validate_plan(plan)
    capture._validate_bindings(PACKET, plan)
    for mode in plan["schedule"]:
        capture._build_descriptor(PACKET, plan, mode)
    capture._write_json(PACKET / "plan.json", plan)
    print("Fixed protocol frozen; no timing child launched.")

if __name__ == "__main__":
    main()
