#!/usr/bin/env python3
"""Exercise the real scoring functions using copies of historical raw inputs.

These synthetic fixtures are verifier tests, never 0723 measurements. Every
mutation lives in an automatically removed temporary directory.
"""
import copy
import hashlib
import importlib.util
import json
from pathlib import Path
import tempfile

P = Path(__file__).resolve().parent
spec = importlib.util.spec_from_file_location("scoring0723", P / "analyze.py")
A = importlib.util.module_from_spec(spec)
spec.loader.exec_module(A)
plan = json.loads((P / "plan.json").read_text())
checks = []


def check(name, condition):
    assert condition, name
    checks.append({"name": name, "pass": True})


check("warm absolute allowance", A.timing_gate("q8", 100, 108, plan)["pass"])
check("warm regression rejected", not A.timing_gate("q8", 100, 120, plan)["pass"])
check("build has no warm allowance", not A.timing_gate("q2", 100, 108, plan)["pass"])
check("workflow regression rejected", not A.timing_gate("open-plus-eight", 100, 108, plan)["pass"])

case = next(c for c in plan["cases"] if c["case"] == "Simple-stored-2097152")
small = copy.deepcopy(plan)
small["cases"] = [case]
small["native"]["groups"] = 2
historical = P.parent / "change-0690"
with tempfile.TemporaryDirectory(prefix="litchi-0723-verifier-") as temporary:
    A.CAPTURES = Path(temporary)
    native = A.CAPTURES / "native"
    native.mkdir()
    for mode in ("owned", "file"):
        original = json.loads((historical / "native" / "baseline" / f"aa1-{case['case']}-{mode}.json").read_text())
        for leg in plan["native"]["legs"]:
            (native / f"{leg}-{case['case']}-{mode}.json").write_text(json.dumps(original))
    failures = []
    rows = A.compare_native(small, failures)
    check("native positive control", not failures and len(rows) == 2 and all(r["timing_gate_pass"] for r in rows))
    changed = native / f"b1-{case['case']}-owned.json"
    original = json.loads(changed.read_text())
    mutation = copy.deepcopy(original)
    mutation["records"][0]["queries"][0]["outcome"]["value"] = {"kind": "String", "value": "corrupt"}
    changed.write_text(json.dumps(mutation))
    failures = []
    rows = A.compare_native(small, failures)
    check("actual outcome mutation rejected", bool(failures) and any(not r["outcome_equal"] for r in rows))
    mutation = copy.deepcopy(original)
    for row in mutation["records"]:
        row["queries"][7]["elapsed_ns"] *= 10
    changed.write_text(json.dumps(mutation))
    failures = []
    rows = A.compare_native(small, failures)
    check("actual timing mutation rejected", any(not r["timing_gate_pass"] for r in rows))

    allocator = A.CAPTURES / "allocator"
    allocator.mkdir()
    for mode in ("owned", "file"):
        for operation in plan["allocator"]["operations"]:
            original = json.loads((historical / "costs" / "baseline" / f"alloc-{case['case']}-{mode}-{operation}-0.json").read_text())
            for phase in ("baseline", "candidate"):
                for repeat in range(plan["allocator"]["repeats"]):
                    (allocator / f"{phase}-{case['case']}-{mode}-{operation}-{repeat}.json").write_text(json.dumps(original))
    failures = []
    rows = A.compare_allocator(small, failures)
    check("allocator positive control", not failures and len(rows) == 8 and all(r["gate_pass"] for r in rows))
    for delta in (1, -1):
        for repeat in range(plan["allocator"]["repeats"]):
            source = allocator / f"baseline-{case['case']}-owned-q3-{repeat}.json"
            target = allocator / f"candidate-{case['case']}-owned-q3-{repeat}.json"
            mutation = json.loads(source.read_text())
            mutation["allocated_bytes"] += delta
            target.write_text(json.dumps(mutation))
        failures = []
        rows = A.compare_allocator(small, failures)
        check(f"strict allocation delta {delta} rejected", bool(failures) and any(not r["gate_pass"] for r in rows))
    # Restore q3 before checking the independent q2 build allowance boundary.
    for repeat in range(plan["allocator"]["repeats"]):
        (allocator / f"candidate-{case['case']}-owned-q3-{repeat}.json").write_bytes(
            (allocator / f"baseline-{case['case']}-owned-q3-{repeat}.json").read_bytes())
    allowance = plan["allocator"]["q2_allowance"]["allocated_bytes"]
    for extra in (allowance, allowance + 1):
        for repeat in range(plan["allocator"]["repeats"]):
            mutation = json.loads((allocator / f"baseline-{case['case']}-owned-q2-{repeat}.json").read_text())
            mutation["allocated_bytes"] += extra
            (allocator / f"candidate-{case['case']}-owned-q2-{repeat}.json").write_text(json.dumps(mutation))
        failures = []
        rows = A.compare_allocator(small, failures)
        check(f"build allowance boundary {extra}", all(r["gate_pass"] for r in rows) == (extra == allowance))

print(json.dumps({"scope": "synthetic verifier checks, not performance evidence", "analysis_sha256": hashlib.sha256((P / "analyze.py").read_bytes()).hexdigest(), "checks": checks}, indent=2))
