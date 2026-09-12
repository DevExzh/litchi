"""Read-only negative custody probes for the 0534 verifier.

The probes mutate deep copies of retained JSON values in memory.  They never
write an evidence report, run Rust, or change the bundle.
"""

from __future__ import annotations

import copy
import datetime as dt
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
    """Build the smallest valid profile shape for the independent gate tests."""
    result = {}
    for stage, multiplier in (("baseline", 10), ("candidate", 5)):
        rows = []
        for repeat in (1, 2):
            for group, shape in (("xls-owned", None), ("cfb-few-large", "few-large")):
                attribution = []
                for _ in range(5):
                    attribution.append({
                        "constructor": {"inclusive_ir": 100 * multiplier},
                        "functions": {"physical_reconciliation": {
                            "target": (
                                "litchi_cfb::file::OleFile<R>::"
                                "validate_physical_sector_layout"
                            ),
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
    results = []
    plan = V.check_plan()
    folder = V.BASELINE
    name = "native-r1-xls"
    receipt_path = folder / (name + ".receipt.json")
    receipt = V.read_json(receipt_path)
    expected = V.expected_artifacts(folder, name, "native")
    V.validate_artifacts(folder, receipt, expected, "positive control")

    for label, mutate in (
        ("omitted-host-artifact", lambda value: value["artifacts"].pop(name + ".host.json")),
        ("forged-raw-vector-hash", lambda value: value["artifacts"].__setitem__(
            name + ".json", "0" * 64)),
        ("unexpected-extra-artifact", lambda value: value["artifacts"].__setitem__(
            "extra.json", "0" * 64)),
    ):
        changed = copy.deepcopy(receipt)
        mutate(changed)
        results.append(must_reject(
            label, lambda changed=changed, label=label: V.validate_artifacts(
                folder, changed, expected, label
            )
        ))

    now = dt.datetime(2026, 9, 12, tzinfo=dt.timezone.utc)
    intervals = [(now, now + dt.timedelta(seconds=2)),
                 (now + dt.timedelta(seconds=1), now + dt.timedelta(seconds=3))]
    results.append(must_reject("overlapping-capture-intervals",
                               lambda: V.check_serial(intervals, "probe")))

    changed_plan = copy.deepcopy(plan)
    changed_plan["review"]["primary_cases"] = changed_plan["review"]["primary_cases"][:1]
    original_read = V.read_json

    def altered_plan(path, *args, **kwargs):
        return changed_plan if path == V.PLAN else original_read(path, *args, **kwargs)

    with patch.object(V, "read_json", side_effect=altered_plan):
        results.append(must_reject("silently-narrowed-primary-gate", V.check_plan))

    profiles = synthetic_profiles()
    physical_rows, physical_gate = V.physical_profile_gate(profiles)
    assert len(physical_rows) == 4 and physical_gate
    forged = copy.deepcopy(profiles)
    forged["candidate"]["profiles"][0]["constructor_attribution"][0]["functions"][
        "physical_reconciliation"
    ]["self_ir"] = 1000
    _, forged_gate = V.physical_profile_gate(forged)
    assert not forged_gate, "physical-Ir increase was accepted"

    quality, quality_intervals = V.validate_quality(plan)
    optional_intervals = V.validate_optional_quality_receipts(plan)
    assert quality["checks"] == len(V.QUALITY_NAMES) == 14
    assert len(quality_intervals) == 14
    assert len(optional_intervals) == sum(
        len(list((V.HERE / stage).glob('check-*.receipt.json')))
        for stage in ('baseline', 'candidate', 'final') if stage != quality['stage']
    ), 'selected quality stage was counted again'
    return {
        "status": "pass",
        "tests": results,
        "quality_checks": quality["checks"],
        "quality_executed_tests": quality["executed_tests"],
        "optional_quality_intervals": len(optional_intervals),
        "physical_gate_rows": len(physical_rows),
        "scope": (
            "Five in-memory custody probes plus physical-Ir and non-duplicated "
            "quality positive controls; no Rust execution or evidence writes."
        ),
    }


if __name__ == "__main__":
    import json

    print(json.dumps(run_tests(), indent=2, sort_keys=True))
