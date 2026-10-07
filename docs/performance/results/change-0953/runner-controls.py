#!/usr/bin/env python3
"""No-workload controls for the 0953 capture runner v5.

The negative controls mutate in-memory JSON.  The positive control constructs
an isolated temporary packet containing only descriptors and summaries: its
profile has 8,879 perf-event rows while its semantic report has 200 measured
iterations.  No control reads or writes the sealed capture tree and none
starts perf or the diagnostic binary.
"""
from __future__ import annotations

import copy
import hashlib
import importlib.util
import json
from pathlib import Path
import stat
import tempfile
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
RUNNER = ROOT / "tools" / "capture_0953_v5.py"
PLAN = ROOT / "plan-0953-v1.json"
OUTPUT = ROOT / "receipts" / "capture-v5-controls-v4.json"


def load_runner() -> Any:
    spec = importlib.util.spec_from_file_location("capture_0953_v5_controls_runner", RUNNER)
    if spec is None or spec.loader is None:
        raise RuntimeError("cannot load capture runner v5")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def sha(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def descriptor(path: Path) -> dict[str, Any]:
    value = path.stat()
    return {"path": str(path.resolve()), "bytes": value.st_size,
            "sha256": sha(path), "mode": stat.S_IMODE(value.st_mode),
            "executable": bool(value.st_mode & stat.S_IXUSR)}


def expect_rejection(function: Any, label: str) -> str:
    try:
        function()
    except Exception as error:  # CaptureError is intentionally imported indirectly.
        return f"{label}: {type(error).__name__}: {error}"
    raise AssertionError(f"{label} was accepted")


def positive_admission_control(runner: Any, source_plan: dict[str, Any],
                               binary: Path) -> None:
    """Exercise the full admission path without pretending an event capture exists."""
    with tempfile.TemporaryDirectory(prefix="litchi-0953-v5-admission-") as raw:
        packet = Path(raw)
        capture = packet / "capture"
        analysis_root = packet / "analysis"
        child = capture / "00-medium"
        child.mkdir(parents=True)
        analysis_root.mkdir(parents=True)
        plan_path = packet / "plan.json"
        plan = copy.deepcopy(source_plan)
        plan["runner"] = descriptor(RUNNER)
        plan["analyzer"] = descriptor(ROOT / "tools" / "analyze_0953_v3.py")
        plan_path.write_text(json.dumps(plan, sort_keys=True) + "\n", encoding="utf-8")
        plan_desc = descriptor(plan_path)
        runner_desc = descriptor(RUNNER)
        binary_desc = descriptor(binary)
        row = plan["schedule"][0]
        child_receipt_path = child / "receipt.json"
        child_receipt = {
            "schema": "litchi.performance.0953.child-receipt.v1",
            "status": "raw_captured", "row": row, "plan": plan_desc,
            "runner": runner_desc, "binary": binary_desc, "retry": False,
        }
        child_receipt_path.write_text(json.dumps(child_receipt, sort_keys=True) + "\n",
                                      encoding="utf-8")
        child_desc = descriptor(child_receipt_path)
        analysis_path = analysis_root / "native-admission-00.json"
        # 8,879 event headers are deliberately distinct from the 200 semantic
        # iterations; the runner must preserve that distinction.
        owner_samples = 100
        owner_period = 100000
        unknown_samples = 8779
        unknown_period = 8779000
        analysis = {
            "schema": "litchi.performance.0953.native-admission.v3",
            "status": "pass", "plan": plan_desc, "runner": runner_desc,
            "binary": binary_desc, "child": child_desc,
            "semantic": {"status": "pass", "shape": "medium",
                          "samples": 200, "warmup": 0},
            "profile": {
                "all": {"samples": 8879, "period": owner_period + unknown_period},
                "owner_qualified": {"samples": owner_samples, "period": owner_period},
                "unknown_or_unqualified": {"samples": unknown_samples,
                                            "period": unknown_period},
                "gate": {"lost_records_zero": True,
                          "fully_resolved_owner_samples_min": True,
                          "owner_period_min": True},
                "admitted": True,
            },
        }
        analysis_path.write_text(json.dumps(analysis, sort_keys=True) + "\n",
                                 encoding="utf-8")
        admission = {
            "schema": "litchi.performance.0953.first-child-admission.v1",
            "status": "pass", "child_index": 0,
            "raw_validation": "pass", "semantic_validation": "pass",
            "retry": False, "replacement": False, "plan": plan_desc,
            "runner": runner_desc, "binary": binary_desc,
            "analyzer": plan["analyzer"], "child_receipt": child_desc,
            "analysis_result": descriptor(analysis_path),
            "owner_samples": owner_samples, "owner_period": owner_period,
            "minimum_owner_samples": 100, "minimum_owner_period": 100000,
        }
        saved = runner.PACKET, runner.CAPTURE_ROOT, runner.ANALYSIS_ROOT
        runner.PACKET, runner.CAPTURE_ROOT, runner.ANALYSIS_ROOT = packet, capture, analysis_root
        try:
            runner._require_first_child_admission_value(admission, plan, plan_path, binary)
        finally:
            runner.PACKET, runner.CAPTURE_ROOT, runner.ANALYSIS_ROOT = saved


def main() -> int:
    runner = load_runner()
    plan = runner.load_json(PLAN)
    schedule_mutation = copy.deepcopy(plan)
    schedule_mutation["schedule"][0]["shape"] = "mutated-schedule"
    schedule_error = expect_rejection(
        lambda: runner.schedule_from_plan(schedule_mutation),
        "schedule mutation")

    # Make a minimal admission object that reaches the descriptor comparison,
    # then point its plan binding at an unrelated real file.  The real
    # first-child admission is never opened or changed by this control.
    binary = Path(plan["binary"]["path"])
    wrong_plan_binding = descriptor(runner.PERF)
    admission = {
        "schema": "litchi.performance.0953.first-child-admission.v1",
        "status": "pass", "child_index": 0,
        "raw_validation": "pass", "semantic_validation": "pass",
        "retry": False, "replacement": False,
        "plan": wrong_plan_binding,
    }
    admission_error = expect_rejection(
        lambda: runner._require_first_child_admission_value(
            admission, plan, PLAN, binary),
        "first-child plan-binding mutation")
    positive_admission_control(runner, plan, binary)

    result = {
        "schema": "litchi.performance.0953.capture-v5-controls.v4",
        "status": "pass",
        "scope": "Synthetic temporary descriptors and in-memory contract controls only; no workload, build, perf capture, decoder export, or sealed-packet mutation.",
        "runner": descriptor(RUNNER),
        "plan": descriptor(PLAN),
        "controls": {
            "schedule_mutation_rejected": True,
            "schedule_mutation_error": schedule_error,
            "first_child_plan_binding_mutation_rejected": True,
            "first_child_plan_binding_error": admission_error,
            "positive_8879_profile_events_200_semantic_iterations": True,
        },
    }
    if OUTPUT.exists():
        raise RuntimeError(f"refusing to overwrite retained control receipt: {OUTPUT}")
    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    OUTPUT.write_text(json.dumps(result, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    print(json.dumps({"status": "pass", "output": str(OUTPUT), "sha256": sha(OUTPUT)},
                     sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
