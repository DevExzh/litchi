"""Read-only positive controls and negative custody probes for 0535.

All mutations stay in deep copies.  This module does not build, capture, or
write either a retained report or any evidence artifact.
"""

from __future__ import annotations

import copy
import datetime as dt
from pathlib import Path
import sys

sys.dont_write_bytecode = True
sys.path.insert(0, str(Path(__file__).resolve().parent))
import verify as V  # noqa: E402


def must_reject(name: str, action) -> dict[str, str]:
    try:
        action()
    except V.VerificationError as error:
        return {"name": name, "status": "pass", "observed_refusal": str(error)}
    raise AssertionError("tampered evidence was accepted: " + name)


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
    changed_plan["performance_claim"] = "a speedup"
    original_read = V.read_json

    def altered_plan(path, *args, **kwargs):
        return changed_plan if path == V.PLAN else original_read(path, *args, **kwargs)

    from unittest.mock import patch
    with patch.object(V, "read_json", side_effect=altered_plan):
        checks.append(must_reject("forged-diagnostic-claim", V.check_plan))

    prior = V.validate_prior_quality(plan)
    assert prior["checks"] == 14 and prior["executed_tests"] == 4382
    builds, _ = V.validate_build(plan)
    binary_sha = builds["binary"]["sha256"]
    numeric = V.validate_numeric(plan, binary_sha)
    instruction = V.validate_instruction_analysis(plan, binary_sha)
    inventory = V.validate_inventory(plan)
    assert numeric["native_rows"] == 4 and instruction["profile_rows"] == 4
    assert inventory["receipts"] == 19
    return {
        "status": "pass",
        "checks": checks,
        "prior_quality_checks": prior["checks"],
        "prior_quality_executed_tests": prior["executed_tests"],
        "native_rows": numeric["native_rows"],
        "instruction_profile_rows": instruction["profile_rows"],
        "receipt_inventory": inventory["receipts"],
        "scope": (
            "Five in-memory custody probes plus prior-quality, numeric, "
            "instruction-analysis, and receipt-inventory positive controls; "
            "no Rust execution or evidence writes."
        ),
    }


if __name__ == "__main__":
    import json

    print(json.dumps(run_tests(), indent=2, sort_keys=True))
