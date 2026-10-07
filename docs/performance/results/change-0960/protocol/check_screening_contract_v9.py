#!/usr/bin/env python3
"""Offline controls for the screen-v2/0960 reducer contract.

These controls exercise plan canonicalization and binding admission only.  They
never launch the screen runner or a benchmark child.
"""
from __future__ import annotations

import copy
import hashlib
import importlib.util
import json
import os
from pathlib import Path

ROOT = Path("/home/zhuhe/litchi-performance-0960").resolve()
PLAN = ROOT / "plan-template-0960-v14.json"
REDUCER = ROOT / "tools/reduce_0960_v3.py"
RUNNER = ROOT / "tools/screen_0960_v7.py"
CONTRACT = ROOT / "review/screening-plan-contract-v2.json"
OUTPUT = ROOT / "review/screening-controls-0960-v9.json"


def descriptor(path: Path, executable: bool | None = None) -> dict:
    stat = path.stat()
    actual = {"path": str(path.resolve()), "bytes": stat.st_size, "mode": stat.st_mode & 0o777, "executable": bool(stat.st_mode & 0o111), "sha256": hashlib.sha256(path.read_bytes()).hexdigest()}
    if executable is not None:
        assert actual["executable"] is executable, f"executable mismatch: {path}"
    return actual


def load(path: Path):
    spec = importlib.util.spec_from_file_location("reduce_0960_v2_controls", path)
    assert spec and spec.loader
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def bound_fixture(module):
    plan = json.loads(PLAN.read_text(encoding="utf-8"))
    plan["status"] = "authorized"
    plan["root_go"] = {"status": "GO", "receipt": descriptor(ROOT / "root-start-v1.json")}
    plan["host"] = {"status": "bound", "machine": "control-only", "receipt": descriptor(ROOT / "root-start-v1.json")}
    plan["measurements_started"] = False
    plan["source"] = copy.deepcopy(plan["source"])
    plan["source"]["baseline_manifest"] = descriptor(ROOT / "source-manifest-baseline-v1.json")
    plan["source"]["candidate_manifest"] = descriptor(ROOT / "source-manifest-baseline-v1.json")
    plan["source"]["baseline_receipt"] = descriptor(ROOT / "receipts/source-custody-v1.json")
    plan["source"]["candidate_freeze"] = descriptor(ROOT / "root-start-v1.json")
    for lane in ("native", "allocator"):
        spec = plan["lanes"][lane]
        spec["candidate_binary"] = copy.deepcopy(spec["baseline_binary"])
        spec["execution"]["workload_cwd"] = str(ROOT)
        plan["lanes"][lane] = spec
    plan["helpers"] = {
        "reducer": descriptor(REDUCER),
        "capture": descriptor(RUNNER, executable=True),
        "controls": descriptor(Path(__file__), executable=True),
    }
    return plan


def reject(name, fn):
    try:
        fn()
    except Exception as exc:
        return {"name": name, "status": "pass", "error": str(exc)}
    return {"name": name, "status": "fail", "error": "invalid binding accepted"}


def main() -> int:
    module = load(REDUCER)
    plan = bound_fixture(module)
    canonical, mapping = module.canonical_plan(plan)
    checks = []
    try:
        module.validate_plan_bindings(plan, canonical, PLAN, REDUCER)
        checks.append({"name": "positive_protocol_plan_binding", "status": "pass", "mapping": mapping})
    except Exception as exc:
        checks.append({"name": "positive_protocol_plan_binding", "status": "fail", "error": str(exc)})
    wrong_binary = copy.deepcopy(plan)
    wrong_binary["lanes"]["native"]["candidate_binary"]["sha256"] = "0" * 64
    checks.append(reject("candidate_binary_digest_mismatch", lambda: module.validate_plan_bindings(wrong_binary, module.canonical_plan(wrong_binary)[0], PLAN, REDUCER)))
    missing_receipt = copy.deepcopy(plan)
    missing_receipt["source"].pop("candidate_freeze")
    checks.append(reject("missing_candidate_freeze", lambda: module.validate_plan_bindings(missing_receipt, module.canonical_plan(missing_receipt)[0], PLAN, REDUCER)))
    wrong_host = copy.deepcopy(plan)
    wrong_host["host"]["receipt"]["sha256"] = "0" * 64
    checks.append(reject("host_receipt_digest_mismatch", lambda: module.validate_plan_bindings(wrong_host, module.canonical_plan(wrong_host)[0], PLAN, REDUCER)))
    wrong_go = copy.deepcopy(plan)
    wrong_go["root_go"]["receipt"]["sha256"] = "0" * 64
    checks.append(reject("root_go_receipt_digest_mismatch", lambda: module.validate_plan_bindings(wrong_go, module.canonical_plan(wrong_go)[0], PLAN, REDUCER)))
    wrong_order = copy.deepcopy(plan)
    wrong_order["comparison"]["blocks"][0]["order"] = ["A", "B", "A", "B"]
    checks.append(reject("native_order_mutation", lambda: module.validate_plan_bindings(wrong_order, module.canonical_plan(wrong_order)[0], PLAN, REDUCER)))
    wrong_runner = copy.deepcopy(plan)
    wrong_runner["helpers"]["capture"]["sha256"] = "0" * 64
    checks.append(reject("runner_digest_mismatch", lambda: module.validate_plan_bindings(wrong_runner, module.canonical_plan(wrong_runner)[0], PLAN, REDUCER)))
    wrong_reducer = copy.deepcopy(plan)
    wrong_reducer["helpers"]["reducer"]["sha256"] = "0" * 64
    checks.append(reject("reducer_digest_mismatch", lambda: module.validate_plan_bindings(wrong_reducer, module.canonical_plan(wrong_reducer)[0], PLAN, REDUCER)))
    result = {
        "schema": "litchi.performance.0960.screening-controls.v3",
        "status": "passed" if all(item["status"] == "pass" for item in checks) else "failed",
        "scope": "protocol-plan canonicalization and binding controls only; no workload or build",
        "contract": descriptor(CONTRACT),
        "plan_fixture": descriptor(PLAN),
        "reducer": descriptor(REDUCER),
        "runner": descriptor(RUNNER, executable=True),
        "checks": checks,
    }
    encoded = (json.dumps(result, indent=2, sort_keys=True) + "\n").encode()
    fd = os.open(OUTPUT, os.O_WRONLY | os.O_CREAT | os.O_EXCL, 0o664)
    with os.fdopen(fd, "wb") as stream:
        stream.write(encoded)
        stream.flush()
        os.fsync(stream.fileno())
    print(json.dumps({"status": result["status"], "output": str(OUTPUT), "bytes": len(encoded)}, sort_keys=True))
    return 0 if result["status"] == "passed" else 1


if __name__ == "__main__":
    raise SystemExit(main())
