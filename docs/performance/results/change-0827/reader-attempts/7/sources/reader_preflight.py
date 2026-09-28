"""Execution-free reader preflight over the retained 0826 report manifest.

This check exercises the report and allocator parsers against historical
schema evidence only.  It does not write a result, start a workload, or reuse
historical timing for the 0827 comparison.
"""

from __future__ import annotations

import copy
import hashlib
import json
from pathlib import Path
import sys
from typing import Any, Callable


PACKET = Path(__file__).resolve().parent
ROOT = PACKET.parents[3]
FIXTURE = ROOT / "docs/performance/results/change-0826/fixtures.json"
sys.path.insert(0, str(PACKET))
import analysis  # noqa: E402
import raw_audit  # noqa: E402

sys.path.insert(0, str(ROOT))
from tools import perf_allocation_schema  # noqa: E402


class PreflightError(AssertionError):
    """A retained reader contract is malformed."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise PreflightError(message)


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def expect_failure(action: Callable[[], Any], label: str) -> None:
    try:
        action()
    except (analysis.ReplayError, raw_audit.AuditError, PreflightError):
        return
    raise PreflightError(f"tamper case was accepted: {label}")


def read_fixture() -> tuple[dict[str, Any], list[dict[str, Any]]]:
    require(FIXTURE.is_file() and not FIXTURE.is_symlink(),
            f"0826 fixture manifest is missing: {FIXTURE}")
    value = json.loads(FIXTURE.read_text(encoding="utf-8"))
    require(value.get("schema") == "litchi.performance.0826.fixtures.v1",
            "0826 fixture schema changed")
    rows = value.get("reports")
    require(isinstance(rows, list) and len(rows) == 325,
            "0826 fixture report count changed")
    require(len({row.get("path") for row in rows}) == len(rows),
            "0826 fixture paths are not unique")
    return value, rows


def check_historical_reports(rows: list[dict[str, Any]]) -> dict[str, Any]:
    counts: dict[str, dict[str, int]] = {}
    cases: set[str] = set()
    observer_example: dict[str, Any] | None = None
    for index, row in enumerate(rows):
        require(isinstance(row, dict), f"fixture row {index} is malformed")
        path = ROOT / row.get("path", "")
        require(path.is_file() and not path.is_symlink(),
                f"fixture report {index} is missing")
        observed = {"path": row["path"], "bytes": path.stat().st_size,
                    "sha256": sha256(path)}
        require(observed == {key: row.get(key) for key in observed},
                f"fixture report {index} identity changed")
        report = json.loads(path.read_text(encoding="utf-8"))
        case = row.get("case")
        mode = row.get("mode")
        require(isinstance(case, str) and mode in {"native", "observer"},
                f"fixture report {index} selector changed")
        perf_allocation_schema.validate_report(
            report,
            expected_case=case,
            expected_samples=row["samples"],
            expected_warmup=row["warmup"],
            mode=mode,
        )
        parsed = raw_audit.parse_report(
            report,
            expected_case=case,
            expected_samples=row["samples"],
            expected_warmup=row["warmup"],
            mode=mode,
            label=f"fixture/{index}",
        )
        require(parsed["case"] == case and len(parsed["samples"]) == row["samples"],
                f"fixture report {index} parser result changed")
        key = f"{row['origin_packet']}/{row['lane']}"
        bucket = counts.setdefault(key, {"reports": 0, "sample_envelopes": 0})
        bucket["reports"] += 1
        bucket["sample_envelopes"] += row["samples"]
        cases.add(case)
        if observer_example is None and mode == "observer":
            observer_example = report
    expected_counts = {
        "0819/qualification": {"reports": 12, "sample_envelopes": 12},
        "0819/native": {"reports": 72, "sample_envelopes": 2160},
        "0819/observer": {"reports": 24, "sample_envelopes": 72},
        "0821/qualification": {"reports": 24, "sample_envelopes": 24},
        "0821/native": {"reports": 144, "sample_envelopes": 4320},
        "0821/observer": {"reports": 48, "sample_envelopes": 144},
        "0825/failed-qualification": {"reports": 1, "sample_envelopes": 1},
    }
    require(counts == expected_counts, "0826 historical lane counts changed")
    require(cases == {case["case"] for case in analysis.CASES},
            "0826 historical case coverage changed")
    require(observer_example is not None, "no observer fixture was retained")
    return {"counts": counts, "cases": len(cases), "observer": observer_example}


def check_signed_allocation(observer: dict[str, Any]) -> None:
    """Exercise signed net-live acceptance and conservation tamper rejection."""
    candidate = copy.deepcopy(observer)
    allocation = candidate["results"][0]["operation_metrics"]["allocation"]
    before = allocation["live_bytes_before"]["values"]
    after = allocation["live_bytes_after"]["values"]
    allocated = allocation["allocated_bytes"]["values"]
    deallocated = allocation["deallocated_bytes"]["values"]
    peak_before = allocation["peak_live_bytes_before"]["values"]
    peak_after = allocation["peak_live_bytes_after"]["values"]
    region_peak = allocation["region_peak_live_bytes"]["values"]
    old_before = before[0]
    before[0] = after[0] + 1
    increase = before[0] - old_before
    deallocated[0] += increase
    peak_before[0] = max(peak_before[0], before[0])
    peak_after[0] = max(peak_after[0], peak_before[0])
    region_peak[0] = max(region_peak[0], before[0])
    parsed = raw_audit.parse_alloc(candidate, len(before), "observer", "signed-net-live")
    require(parsed is not None and parsed["net_live"][0] < 0,
            "signed net-live vector was not retained")

    tampered = copy.deepcopy(candidate)
    tampered["results"][0]["operation_metrics"]["allocation"][
        "live_bytes_after"]["values"][0] += 1
    expect_failure(
        lambda: raw_audit.parse_alloc(tampered, len(before), "observer", "bad-conservation"),
        "allocation live-byte conservation",
    )


def check_canonical_and_retry_helpers() -> None:
    tuple_value = {"cases": [{"changed_members": ("a.xml", "b.xml")}]}
    list_value = {"cases": [{"changed_members": ["a.xml", "b.xml"]}]}
    require(analysis.canonical_json(tuple_value) == analysis.canonical_json(list_value),
            "canonical JSON did not normalize tuple/list evidence")
    tampered = {"cases": [{"changed_members": ["a.xml", "c.xml"]}]}
    require(analysis.canonical_json(list_value) != analysis.canonical_json(tampered),
            "canonical JSON accepted changed preservation evidence")

    accepted = {"admission-before-retry2/artifact-audit-receipt.json": {}}
    require(analysis.accepted_admission_attempt_root("before", accepted) ==
            PACKET / "admission-before-retry2",
            "accepted admission retry was not resolved")
    expect_failure(
        lambda: analysis.accepted_admission_attempt_root(
            "before", {"admission-after-retry2/artifact-audit-receipt.json": {}}),
        "wrong-leg admission retry",
    )
    expect_failure(
        lambda: analysis.accepted_admission_attempt_root(
            "before", {"admission-before/a.json": {}, "admission-before-retry1/b.json": {}}),
        "multiple accepted admission roots",
    )
    expect_failure(
        lambda: analysis.accepted_admission_attempt_root("before", {"../escape.json": {}}),
        "escaped admission path",
    )


def main(argv: list[str] | None = None) -> int:
    args = sys.argv[1:] if argv is None else argv
    require(args in ([], ["--write"]),
            "reader preflight accepts no arguments or the suite's --write marker")
    try:
        _, rows = read_fixture()
        value = check_historical_reports(rows)
        check_signed_allocation(value["observer"])
        check_canonical_and_retry_helpers()
        print("0827 reader preflight PASS: 325 historical reports; signed allocation and retry tamper checks")
        return 0
    except (PreflightError, raw_audit.AuditError, analysis.ReplayError,
            OSError, ValueError, KeyError, TypeError, IndexError) as error:
        print(f"0827 reader preflight failed: {error}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
