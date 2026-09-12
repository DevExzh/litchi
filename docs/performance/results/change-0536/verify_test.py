"""Read-only custody probes for the 0536 matched verifier.

Mutations are limited to deep copies of retained JSON values.  The probes do
not build, capture, run Rust, or write an evidence report.
"""

from __future__ import annotations

import copy
import datetime as dt
import json
from pathlib import Path
import sys
from unittest.mock import patch

sys.dont_write_bytecode = True
sys.path.insert(0, str(Path(__file__).resolve().parent))
import verify as V  # noqa: E402


def must_reject(name: str, action) -> dict[str, str]:
    try:
        action()
    except V.VerificationError as error:
        return {"name": name, "status": "pass", "observed_refusal": str(error)}
    raise AssertionError("tampered evidence was accepted: " + name)


def synthetic_profiles() -> dict[str, dict[str, list[dict]]]:
    """Build the smallest profile matrix accepted by both independent gates."""
    result = {}
    for stage, multiplier in (("baseline", 10), ("candidate", 5)):
        rows = []
        for repeat in (1, 2):
            for group, shape in (("xls-owned", None), ("cfb-few-large", "few-large")):
                attribution = []
                for _ in range(5):
                    attribution.append({
                        "constructor": {"inclusive_ir": 100 * multiplier},
                        "functions": {"collector": {
                            "target": V.COLLECTOR_TARGET,
                            "self_ir": multiplier,
                            "out_of_line": True,
                            "inlined_or_absent": False,
                            "incoming_edge_count": 1,
                        }},
                    })
                rows.append({"group": group, "shape": shape, "repeat": repeat,
                             "constructor_attribution": attribution})
        result[stage] = {"profiles": rows}
    return result


def run_tests() -> dict:
    plan = V.check_plan()
    folder = V.BASELINE
    name = "native-r1-xls"
    receipt = V.read_json(folder / (name + ".receipt.json"))
    expected = V.expected_artifacts(folder, name, "native")
    V.validate_artifacts(folder, receipt, expected, "positive control")

    checks = []
    for label, mutate in (
        ("omitted-host-artifact", lambda value: value["artifacts"].pop(name + ".host.json")),
        ("forged-raw-vector-hash", lambda value: value["artifacts"].__setitem__(
            name + ".json", "0" * 64)),
        ("unexpected-extra-artifact", lambda value: value["artifacts"].__setitem__(
            "extra.json", "0" * 64)),
    ):
        changed = copy.deepcopy(receipt)
        mutate(changed)
        checks.append(must_reject(
            label,
            lambda changed=changed, label=label: V.validate_artifacts(
                folder, changed, expected, label
            ),
        ))

    now = dt.datetime(2026, 9, 12, tzinfo=dt.timezone.utc)
    checks.append(must_reject(
        "overlapping-capture-intervals",
        lambda: V.check_serial(
            [(now, now + dt.timedelta(seconds=2)),
             (now + dt.timedelta(seconds=1), now + dt.timedelta(seconds=3))],
            "probe",
        ),
    ))

    changed_plan = copy.deepcopy(plan)
    changed_plan["review"]["primary_cases"] = changed_plan["review"]["primary_cases"][:1]
    original_read = V.read_json

    def altered_plan(path, *args, **kwargs):
        return changed_plan if path == V.PLAN else original_read(path, *args, **kwargs)

    with patch.object(V, "read_json", side_effect=altered_plan):
        checks.append(must_reject("silently-narrowed-primary-gate", V.check_plan))

    profiles = synthetic_profiles()
    constructor_rows, constructor_gate = V.profile_ir_gate(profiles)
    collector_rows, collector_gate = V.collector_profile_gate(profiles)
    assert len(constructor_rows) == 2 and constructor_gate
    assert len(collector_rows) == 4 and collector_gate
    forged = copy.deepcopy(profiles)
    forged["candidate"]["profiles"][0]["constructor_attribution"][0]["functions"][
        "collector"
    ]["self_ir"] = 1000
    _, forged_gate = V.collector_profile_gate(forged)
    assert not forged_gate, "collector-Ir increase was accepted"

    # These positive controls are intentionally run only after the coordinator
    # has retained the complete matrix and quality summary.
    quality, quality_intervals = V.validate_quality(plan)
    optional_intervals = V.validate_optional_quality_receipts(plan, quality["stage"])
    instruction = V.validate_instruction(plan)
    assert quality["checks"] == len(V.QUALITY_NAMES) == 14
    assert len(quality_intervals) == 14
    inventory = V.validate_receipt_inventory(plan)
    timeline = V.validate_receipt_timeline()
    assert inventory["successful_receipts"] >= 30
    assert timeline["native_abba"] == [
        "baseline/native-r1-xls", "baseline/native-r1-cfb",
        "candidate/native-r1-xls", "candidate/native-r1-cfb",
        "candidate/native-r2-cfb", "candidate/native-r2-xls",
        "baseline/native-r2-cfb", "baseline/native-r2-xls",
    ]
    return {
        "status": "pass",
        "checks": checks,
        "quality_checks": quality["checks"],
        "quality_executed_tests": quality["executed_tests"],
        "optional_quality_intervals": len(optional_intervals),
        "constructor_gate_rows": len(constructor_rows),
        "collector_gate_rows": len(collector_rows),
        "instruction_reports": len(instruction["instruction_sha256"]),
        "receipt_inventory": inventory["successful_receipts"],
        "serial_intervals": timeline["serial_intervals"],
        "scope": (
            "In-memory receipt/plan probes plus constructor/collector-Ir, quality, "
            "inventory and ABBA positive controls; no Rust execution or evidence writes."
        ),
    }


if __name__ == "__main__":
    print(json.dumps(run_tests(), indent=2, sort_keys=True))
