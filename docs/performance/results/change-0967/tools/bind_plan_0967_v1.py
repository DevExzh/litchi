#!/usr/bin/env python3
"""Bind built executable hashes into a fresh 0967 plan copy.

This tool does not build or launch anything. It writes a new plan and leaves the
prospective input untouched. Capture remains disabled until root deliberately
changes the copied plan status to ``capture_authorized`` after host/GO review.
"""
from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path
from typing import Any

ROLES = ("native_baseline", "native_candidate", "allocator_baseline", "allocator_candidate", "control_baseline", "control_candidate", "timed_control_native_baseline", "timed_control_native_candidate", "timed_control_allocator_baseline", "timed_control_allocator_candidate")

def sha(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()

def load(path: Path) -> dict[str, Any]:
    value = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(value, dict):
        raise SystemExit(f"{path}: expected JSON object")
    return value

def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--plan", type=Path, default=Path("/home/zhuhe/litchi-performance-0967/plans/plan-0967-v1.json"))
    ap.add_argument("--output", type=Path, default=Path("/home/zhuhe/litchi-performance-0967/plans/plan-0967-bound-v1.json"))
    ap.add_argument("--binary", action="append", default=[], metavar="ROLE=PATH")
    ap.add_argument("--control-receipt", type=Path)
    args = ap.parse_args()
    plan_path = args.plan.resolve()
    plan = load(plan_path)
    if plan.get("schema") != "litchi.performance.0967.plan.v1":
        raise SystemExit("unexpected plan schema")
    bindings = {}
    for raw in args.binary:
        try:
            role, raw_path = raw.split("=", 1)
        except ValueError as exc:
            raise SystemExit(f"binary binding must be ROLE=PATH: {raw}") from exc
        if role not in ROLES:
            raise SystemExit(f"unknown binary role {role}")
        path = Path(raw_path).resolve()
        if not path.is_file():
            raise SystemExit(f"missing binary {path}")
        expected = Path(plan["binaries"][role]["path"]).resolve()
        if path != expected:
            raise SystemExit(f"{role}: path differs from plan: {path} != {expected}")
        bindings[role] = {"path": str(path), "sha256": sha(path)}
    if set(bindings) != set(ROLES):
        raise SystemExit(f"all six binary roles must be bound; missing {sorted(set(ROLES) - set(bindings))}")
    for role, row in bindings.items():
        plan["binaries"][role].update(row)
    if args.control_receipt:
        receipt = args.control_receipt.resolve()
        doc = load(receipt)
        if doc.get("status") not in {"passed", "sealed"}:
            raise SystemExit("control receipt is not sealed/passed")
        plan["controls"]["status"] = "bound"
        plan["controls"]["probe_receipt"] = {"path": str(receipt), "sha256": sha(receipt), "status": doc.get("status")}
    plan["parent_plan"] = {"path": str(plan_path), "sha256": sha(plan_path)}
    plan["status"] = "pending_capture_authorization"
    plan["build_policy"]["bound_binaries"] = True
    # The timing-capable control binding is intentionally independent of the
    # correctness probe and must be supplied in a reviewed plan update.
    plan["binding_note"] = "Executable hashes are bound. Root must bind timed positive/dense/false controls, host state, and GO before setting capture_authorized."
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(plan, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({"status": plan["status"], "output": str(args.output), "sha256": sha(args.output)}))
    return 0

if __name__ == "__main__":
    raise SystemExit(main())
