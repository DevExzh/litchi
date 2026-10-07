#!/usr/bin/env python3
"""Independent schedule control for the corrected 0948 capture helper.

The checker calls the helper's actual expected_jobs/stage_counts functions and
compares their complete order to the frozen 16/48/24 schedule.  It never
launches a child, build, or workload.  The final invocation must name the
final helper and final plan explicitly; no draft plan is accepted.
"""
from __future__ import annotations

import argparse
import hashlib
import importlib.util
import json
import os
from pathlib import Path
from typing import Any

EXPECTED = {
    "qualification": {"children": 16, "samples": 2, "warmup": 0},
    "native": {"children": 48, "samples": 100, "warmup": 10},
    "allocator": {"children": 24, "samples": 5, "warmup": 2},
}
ROLES = ("baseline", "candidate")
CASES = (
    "primary-percent-medium", "primary-percent-dense-sparse",
    "control-xlsx-one-edit-medium", "control-xlsx-managed-one-edit-medium",
)


def sha(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def load(path: Path) -> Any:
    spec = importlib.util.spec_from_file_location("capture_0948_schedule_control", path)
    if not spec or not spec.loader:
        raise RuntimeError(f"cannot import {path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def require(condition: bool, message: str) -> None:
    if not condition:
        raise AssertionError(message)


def check_stage(capture: Any, plan: dict[str, Any], stage: str) -> dict[str, Any]:
    jobs = capture.expected_jobs(plan, stage)
    wanted = EXPECTED[stage]
    require(len(jobs) == wanted["children"], f"{stage}: {len(jobs)} != {wanted['children']}")
    count, samples, warmup = capture.stage_counts(plan, stage)
    require((count, samples, warmup) ==
            (wanted["children"], wanted["samples"], wanted["warmup"]),
            f"{stage}: stage_counts {(count, samples, warmup)} != expected")
    if stage == "qualification":
        expected = [(kind, role, case) for kind in ("native", "allocator")
                    for role in ROLES for case in CASES]
        actual = [(job["kind"], job["role"], job["case"]) for job in jobs]
        require(actual == expected, "qualification order/roles/cases changed")
    else:
        repeats = 6 if stage == "native" else 3
        kind = "native" if stage == "native" else "allocator"
        width = len(CASES) * len(ROLES)
        require(len(jobs) == repeats * width, f"{stage}: repeat/block width changed")
        for index in range(repeats):
            block = jobs[index * width:(index + 1) * width]
            expected_roles = ROLES if index % 2 == 0 else tuple(reversed(ROLES))
            expected_cases = CASES if index % 2 == 0 else tuple(reversed(CASES))
            expected_block = [(kind, case, role)
                             for case in expected_cases for role in expected_roles]
            actual_block = [(job["kind"], job["case"], job["role"]) for job in block]
            require(actual_block == expected_block,
                    f"{stage} index {index} role/case alternation changed")
            require({job["index"] for job in block} == {index},
                    f"{stage} index {index} labels changed")
    return {"children": len(jobs), "samples": samples, "warmup": warmup,
            "first": jobs[0], "last": jobs[-1]}


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--capture", type=Path, required=True)
    parser.add_argument("--plan", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    capture_path = args.capture.resolve()
    plan_path = args.plan.resolve()
    out = args.output.resolve()
    plan = json.loads(plan_path.read_text(encoding="utf-8"))
    require(plan.get("status") in {"frozen", "capture-ready", "ready"},
            f"plan is not final/capture-ready: {plan.get('status')!r}")
    capture = load(capture_path)
    checks: list[dict[str, Any]] = []
    for stage in ("qualification", "native", "allocator"):
        summary = check_stage(capture, plan, stage)
        checks.append({"id": f"{stage}_expected_jobs_and_stage_counts", "status": "pass",
                       **{key: summary[key] for key in ("children", "samples", "warmup")}})
    require(plan.get("workload_cwd") is not None, "final plan lacks workload_cwd")
    cwd = capture.workload_cwd(plan)
    checks.append({"id": "workload_cwd_is_real_directory", "status": "pass", "path": str(cwd)})

    mutated = json.loads(json.dumps(plan))
    mutated["allocator"]["repeats"] = 2
    try:
        capture.allocator_settings(mutated)
    except Exception as error:
        checks.append({"id": "wrong_allocator_repeat_count_rejected", "status": "pass",
                       "error": type(error).__name__ + ": " + str(error)})
    else:
        checks.append({"id": "wrong_allocator_repeat_count_rejected", "status": "fail"})

    receipt = {
        "schema": "litchi.performance.0948.schedule-control.v1",
        "status": "pass" if all(item["status"] == "pass" for item in checks) else "fail",
        "scope": "actual helper expected_jobs/stage_counts only; no child/build/workload execution",
        "capture": {"path": str(capture_path), "bytes": capture_path.stat().st_size,
                     "sha256": sha(capture_path)},
        "plan": {"path": str(plan_path), "bytes": plan_path.stat().st_size,
                 "sha256": sha(plan_path), "status": plan.get("status")},
        "checks": checks,
    }
    data = (json.dumps(receipt, indent=2, sort_keys=True, allow_nan=False) + "\n").encode()
    if out.exists():
        raise RuntimeError(f"refusing to overwrite {out}")
    with out.open("xb") as stream:
        stream.write(data)
        stream.flush()
        os.fsync(stream.fileno())
    print(json.dumps({"status": receipt["status"], "output": str(out),
                      "bytes": len(data), "sha256": sha(out)}, sort_keys=True))


if __name__ == "__main__":
    main()
