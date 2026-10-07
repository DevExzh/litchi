#!/usr/bin/env python3
"""Independent offline checks for the 0948 draft reducer.

This file never starts a workload or build and does not modify the packet.
It exercises the imported reducer's paired-gate functions with small
synthetic rows, so protocol mistakes are distinguished from missing capture
data.  The input tool and draft plan are hashed in the output receipt.
"""

from __future__ import annotations

import hashlib
import importlib.util
import json
import copy
from pathlib import Path
import shutil
import tempfile
from typing import Any


PACKET = Path("/home/zhuhe/litchi-performance-0948")
REDUCER_PATH = PACKET / "tools" / "reduce_0948_v1.py"
PLAN_PATH = PACKET / "plan-v1-draft.json"
OUT = PACKET / "review" / "protocol" / "review_protocol_0948_v1.json"


def sha256(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def load_reducer() -> Any:
    spec = importlib.util.spec_from_file_location("reducer_0948_review_v1", REDUCER_PATH)
    assert spec and spec.loader
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def plan() -> dict[str, Any]:
    return json.loads(PLAN_PATH.read_text(encoding="utf-8"))


def native_rows(reducer: Any) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    cases = [item[0] for item in reducer.CASES]
    for block in range(6):
        for case_id in cases:
            for role in ("baseline", "candidate"):
                publication = 100
                # Five blocks show a 6% publication improvement, while one
                # block is equal.  The aggregate primary median/bootstrap is
                # still below 1.0, but the equal block is within the approved
                # per-block <=1.05 safety bound.
                if case_id in reducer.CASE_ORDER[:2] and role == "candidate" and block < 5:
                    publication = 94
                rows.append({
                    "stage": "native", "index": block, "role": role,
                    "case_id": case_id, "label": f"{block}/{role}/{case_id}",
                    "whole_workflow_p50": 100,
                    "p95": 100, "p99": 100,
                    "publication_p50": publication,
                    "publication_p95": 100, "publication_p99": 100,
                })
    return rows


def allocator_rows(reducer: Any) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    cases = [item[0] for item in reducer.CASES]
    fields = reducer.ALLOCATOR_SAMPLE_FIELDS
    for repeat in range(3):
        for case_id in cases:
            for role in ("baseline", "candidate"):
                value = 100 if role == "baseline" else 104
                samples = []
                for _ in range(5):
                    sample = {field: value for field in fields}
                    sample.update({"status": "measured", "scope": reducer.ALLOCATOR_SCOPE})
                    samples.append(sample)
                rows.append({
                    "stage": "allocator", "index": repeat, "role": role,
                    "case_id": case_id, "label": f"{repeat}/{role}/{case_id}",
                    "allocator_metrics": {"publication_allocation_metrics": samples},
                    "max_rss_kib": value,
                })
    return rows


def main() -> None:
    reducer = load_reducer()
    frozen_plan = plan()
    # The draft keeps bootstrap at the top level while capture/reduce load it
    # from plan["gate"]; use an explicit in-memory normalization solely to
    # reach the later synthetic gate checks.  The mismatch is itself a
    # pre-capture finding and no packet plan is rewritten.
    working_plan = copy.deepcopy(frozen_plan)
    if "bootstrap" not in working_plan.get("gate", {}):
        working_plan.setdefault("gate", {})["bootstrap"] = working_plan.get("bootstrap")
    checks: list[dict[str, Any]] = []

    native = reducer.native_reduce(native_rows(reducer), working_plan)
    # The approved protocol accepts these rows: primary aggregate estimate is
    # below .95 with upper <1, while every block is <=1.05.  This fixture
    # guards against accidentally moving the aggregate .95 threshold into the
    # per-block safety gate.
    checks.append({
        "id": "primary_block_safety_threshold_is_not_aggregate_threshold",
        "status": "pass" if native["all_case_latency_gate_pass"] else "fail",
        "expected_protocol": "pass",
        "reducer_result": native["all_case_latency_gate_pass"],
        "review_flags": native["review_flags"],
    })

    allocation = reducer.allocator_reduce(allocator_rows(reducer), working_plan)
    checks.append({
        "id": "allocator_five_metric_pair_and_aggregate_safety",
        "status": "pass" if allocation["all_pair_gate_pass"]
        and allocation["all_row_safety_gate_pass"] else "fail",
        "all_pair_gate_pass": allocation["all_pair_gate_pass"],
        "all_row_safety_gate_pass": allocation["all_row_safety_gate_pass"],
    })

    zero_rows = allocator_rows(reducer)
    for row in zero_rows:
        if row["role"] == "baseline":
            for sample in row["allocator_metrics"]["publication_allocation_metrics"]:
                for field in reducer.ALLOCATOR_SAMPLE_FIELDS:
                    sample[field] = 0
            row["max_rss_kib"] = 0
            break
    for row in zero_rows:
        if row["role"] == "candidate":
            row["allocator_metrics"]["publication_allocation_metrics"][0]["allocated_bytes"] = 1
            break
    try:
        reducer.allocator_reduce(zero_rows, working_plan)
    except Exception as error:  # fail closed is the expected result
        checks.append({"id": "nonzero_over_zero_baseline_rejected", "status": "pass",
                       "error": type(error).__name__ + ": " + str(error)})
    else:
        checks.append({"id": "nonzero_over_zero_baseline_rejected", "status": "fail"})

    missing_rows = allocator_rows(reducer)
    del missing_rows[-1]
    try:
        reducer.allocator_reduce(missing_rows, working_plan)
    except Exception as error:
        checks.append({"id": "missing_allocator_child_rejected", "status": "pass",
                       "error": type(error).__name__ + ": " + str(error)})
    else:
        checks.append({"id": "missing_allocator_child_rejected", "status": "fail"})

    missing_metric_rows = allocator_rows(reducer)
    del missing_metric_rows[0]["allocator_metrics"]["publication_allocation_metrics"][0]["allocated_bytes"]
    try:
        reducer.allocator_reduce(missing_metric_rows, working_plan)
    except Exception as error:
        checks.append({"id": "missing_allocator_metric_rejected", "status": "pass",
                       "error": type(error).__name__ + ": " + str(error)})
    else:
        checks.append({"id": "missing_allocator_metric_rejected", "status": "fail"})

    checks.append({"id": "equal_zero_is_explicit_degenerate",
                   "status": "pass" if reducer.ratio(0, 0, "review") == 1.0
                   and reducer.ratio_status(0, 0, "review") == "degenerate_equal_zero"
                   else "fail"})

    with tempfile.TemporaryDirectory(prefix="review-0948-", dir="/home/zhuhe/litchi-performance-0948/review/protocol") as temp:
        moved = Path(temp) / "moved-reducer.py"
        shutil.copy2(REDUCER_PATH, moved)
        checks.append({"id": "tool_copy_hash_stable", "status": "pass"
                       if sha256(moved) == sha256(REDUCER_PATH) else "fail"})

    findings = [
        {"id": "portable-validator-import", "severity": "high",
         "detail": "load_validator passes a packet-relative digest path to spec_from_file_location; running reduction outside packet cwd fails. Output rows also retain absolute folder paths."},
        {"id": "source-end-census", "severity": "medium",
         "detail": "capture child source_end uses source_check(full=False), which checks only candidate files and does not reject an unchanged-file mutation or extra source-root file until the final stage census."},
        {"id": "stale-reuse-helper", "severity": "blocker",
         "detail": "reuse_0948_v1.py still validates 0945 schemas/commit/paths; it cannot validate the 0948 baseline-reuse-v1 receipt or final 0948 build receipts without an explicit adapter."},
        {"id": "packet-tree-symlink-and-extra-census", "severity": "medium",
         "detail": "reducer source_check and child-folder review hash listed files but do not reject extra source files or symlinked role/child directories; the frozen protocol requires a symlink-free exact closure."},
    ]
    if "bootstrap" not in frozen_plan.get("gate", {}):
        findings.insert(0, {"id": "bootstrap-location", "severity": "blocker",
                             "detail": "plan template stores bootstrap outside gate while capture/reduce require gate.bootstrap; freeze must bind one consistent location."})
    if frozen_plan.get("status") not in {"frozen", "capture-ready", "ready"}:
        findings.insert(0, {"id": "plan-not-frozen", "severity": "blocker",
                             "detail": "the inspected input is still a draft and cannot authorize a workload; rerun this receipt against final plan-v2 and helpers."})

    receipt = {
        "schema": "litchi.performance.0948.independent-protocol-review.v1",
        "status": "blocked_before_capture",
        "workloads_started": False,
        "builds_started": False,
        "plan": {"path": str(PLAN_PATH), "bytes": PLAN_PATH.stat().st_size,
                 "sha256": sha256(PLAN_PATH), "status": frozen_plan.get("status")},
        "reducer": {"path": str(REDUCER_PATH), "bytes": REDUCER_PATH.stat().st_size,
                    "sha256": sha256(REDUCER_PATH)},
        "checks": checks,
        "findings": findings,
    }
    OUT.write_text(json.dumps(receipt, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    print(json.dumps({"status": receipt["status"], "output": str(OUT),
                      "sha256": sha256(OUT), "checks": checks}, sort_keys=True))


if __name__ == "__main__":
    main()
