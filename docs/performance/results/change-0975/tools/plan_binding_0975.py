#!/usr/bin/env python3
"""Small 0975 custody checker for the reused 0974 v5 protocol."""
from __future__ import annotations

import hashlib
import json
from pathlib import Path
from typing import Any

ROOT = Path("/home/zhuhe/litchi-performance-0975")
HELPER_SCHEMA = "litchi.performance.0974.helper_seal.v3"
PLAN_SCHEMA = "litchi.performance.0974.plan.v3"
V5_PLAN = Path("/home/zhuhe/litchi-performance-0974/plans/plan-0974-pre-capture-v5.json")
V5_HELPER_SEAL = Path("/home/zhuhe/litchi-performance-0974/receipts/helper-seal-0974-v5.json")
V5_ADAPTER = Path("/home/zhuhe/litchi-performance-0974/tools/protocol_adapter_0974_v5.py")
V3_ADAPTER = Path("/home/zhuhe/litchi-performance-0974/tools/protocol_adapter_0974_v3.py")
V1_ADAPTER = Path("/home/zhuhe/litchi-performance-0974/tools/protocol_adapter_0974_v1.py")
V5_RUNNER = Path("/home/zhuhe/litchi-performance-0974/tools/run_0974_v5.py")
V4_REDUCER = Path("/home/zhuhe/litchi-performance-0974/tools/reduce_0974_v4.py")
V5_GENERATOR = Path("/home/zhuhe/litchi-performance-0974/tools/prepare_plan_0974_v5.py")


class BindingError(RuntimeError):
    pass


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def binding(path: str | Path) -> dict[str, Any]:
    path = Path(path).absolute()
    if not path.is_file():
        raise BindingError(f"missing binding path: {path}")
    return {"path": str(path), "sha256": sha256(path), "bytes": path.stat().st_size}


def same_file_binding(row: Any, path: str | Path) -> bool:
    """Accept retained descriptors that carry extra stable metadata."""
    if not isinstance(row, dict) or not isinstance(path, (str, Path)):
        return False
    expected = binding(path)
    return all(row.get(key) == expected[key] for key in ("path", "sha256", "bytes"))


def load(path: str | Path) -> dict[str, Any]:
    path = Path(path).resolve()
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise BindingError(f"cannot load JSON {path}: {exc}") from exc
    if not isinstance(value, dict):
        raise BindingError(f"JSON root is not an object: {path}")
    return value


def expected_dependencies() -> dict[str, dict[str, Any]]:
    return {
        "reused_adapter_v5": binding(V5_ADAPTER),
        "reused_adapter_v3": binding(V3_ADAPTER),
        "reused_adapter_v1": binding(V1_ADAPTER),
        "reused_runner_v5": binding(V5_RUNNER),
        "reused_reducer_v4": binding(V4_REDUCER),
        "reused_generator_v5": binding(V5_GENERATOR),
        "reused_helper_seal_v5": binding(V5_HELPER_SEAL),
        "reused_plan_v5": binding(V5_PLAN),
        "plan_binding_0975": binding(Path(__file__)),
    }


def check(plan_path: str | Path) -> dict[str, Any]:
    plan_path = Path(plan_path).resolve()
    plan = load(plan_path)
    if plan.get("schema") != PLAN_SCHEMA or plan.get("status") != "pre_capture":
        raise BindingError("0975 plan schema/status mismatch")
    if plan.get("plan_path") != str(plan_path):
        raise BindingError("0975 plan self-binding mismatch")
    if plan.get("outputs") != {
        "capture_root": str((ROOT / "capture/instruction-profile-v1").resolve()),
        "analysis_root": str((ROOT / "analysis").resolve()),
    }:
        raise BindingError("0975 output roots are not exclusive")
    if plan.get("no_build_or_workload") is not True:
        raise BindingError("0975 plan must remain pre-capture")
    batch = plan.get("batch_0975")
    expected_batch = {
        "fresh_cohort": True,
        "not_a_retry_of": "/home/zhuhe/litchi-performance-0974",
        "children": 12,
        "serial": True,
        "cpu": 31,
        "event": "instructions:u",
        "period": 1000000,
        "callgraph": "dwarf,32768",
        "record_mode": ["--per-thread"],
        "record_memory": ["-m", "16M"],
        "record_memory_bytes": 16777216,
        "warmups": 5,
        "samples": 80,
        "minimum_marker_samples_each_child": 100,
        "marker_period_share_positive_each_child": True,
        "required_raw_lost_sum": 0,
        "required_terminal_lost_samples": 0,
        "no_retry": True,
        "stop_on_first_failure": True,
    }
    if batch != expected_batch:
        raise BindingError("0975 fixed batch contract drift")
    expected_deps = expected_dependencies()
    seal_row = plan.get("helper_seal")
    if not isinstance(seal_row, dict):
        raise BindingError("0975 helper seal missing")
    seal_path = Path(seal_row["path"]).resolve()
    if seal_row != binding(seal_path):
        raise BindingError("0975 helper seal binding mismatch")
    seal = load(seal_path)
    if seal.get("schema") != HELPER_SCHEMA or seal.get("status") != "sealed":
        raise BindingError("0975 helper seal schema/status mismatch")
    if seal_row != binding(seal_path):
        raise BindingError("0975 helper seal bytes changed")
    if plan.get("helpers") != seal.get("helpers"):
        raise BindingError("0975 plan/helper map mismatch")
    if plan.get("dependency_bindings") != expected_deps or seal.get("dependency_bindings") != expected_deps:
        raise BindingError("0975 dependency bindings are incomplete or stale")
    for name, row in expected_deps.items():
        if binding(row["path"]) != row:
            raise BindingError(f"0975 dependency changed: {name}")
    for name, row in (seal.get("helpers") or {}).items():
        if binding(row["path"]) != row:
            raise BindingError(f"0975 helper changed: {name}")
    design = plan.get("design")
    if not isinstance(design, dict) or design.get("path") != str((ROOT / "design-0975-v1.md").resolve()) or design != binding(design["path"]):
        raise BindingError("0975 design binding mismatch")
    binary = plan.get("binary")
    fixture = plan.get("fixture")
    if not isinstance(binary, dict) or not isinstance(fixture, dict) or not same_file_binding(binary, binary.get("path")) or not same_file_binding(fixture, fixture.get("path")):
        raise BindingError("0975 binary/fixture custody mismatch")
    if plan.get("perf", {}).get("record_mode", {}).get("tokens") != ["--per-thread"] or plan.get("perf", {}).get("record_memory", {}).get("tokens") != ["-m", "16M"]:
        raise BindingError("0975 native protocol drift")
    return {"status": "pass", "plan": str(plan_path), "plan_sha256": sha256(plan_path), "helper_seal": binding(seal_path), "dependency_count": len(expected_deps)}


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--plan", type=Path, required=True)
    args = parser.parse_args()
    print(json.dumps(check(args.plan), indent=2, sort_keys=True))
