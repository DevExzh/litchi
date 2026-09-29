"""Mutation tests for the independent 0840 evidence reader.

The script copies retained reports into the owned scratch directory, corrupts
one fact at a time, and requires the reader to reject each corruption.  The
retained reports and CFB artifacts are never modified.
"""

from __future__ import annotations

import copy
import json
from pathlib import Path

import driver as d
import readers as r


P = d.P
RUNS = P / "runs"


def _first(lane: str) -> tuple[Path, dict, dict]:
    reports = sorted((RUNS / lane).glob("*/report.json"))
    if not reports:
        raise AssertionError(f"reader mutation tests require one {lane} report")
    report = next(path for path in reports if d.read(path.parent / "started.json")["case"] == "cfb-large")
    folder = report.parent
    started = d.read(folder / "started.json")
    return report, started, d.read(report)


def _reject(name: str, value: dict, original: Path, started: dict,
            *, observer: bool = False, allocation: bool = False,
            passed: list[str]) -> None:
    target = d.SCRATCH / "reader-mutation-0840.json"
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_text(json.dumps(value), encoding="utf-8")
    plan = d.read(P / "plan.json")
    lane = started["lane"]
    config = plan["qualification" if lane == "allocation-preflight" else lane]
    try:
        r.validate_report(target, started["case"], config["samples"], config["warmup"],
                          observer=observer, allocation=allocation)
    except r.ReaderError:
        passed.append(name)
    else:
        raise AssertionError(f"reader accepted mutation {name} from {original}")
    finally:
        target.unlink(missing_ok=True)


def _timed_tests(passed: list[str]) -> None:
    report, started, original = _first("qualification")
    value = copy.deepcopy(original)
    value["schema"] = "wrong"
    _reject("timed-schema", value, report, started, passed=passed)

    value = copy.deepcopy(original)
    value["input"]["members"][0]["sha256"] = "0" * 64
    _reject("timed-input-member-hash", value, report, started, passed=passed)

    value = copy.deepcopy(original)
    value["samples"][0]["wall_ns"] = 0
    _reject("timed-nonpositive-wall", value, report, started, passed=passed)

    value = copy.deepcopy(original)
    value["final_verification"]["output_sha256"] = "0" * 64
    _reject("timed-final-output-hash", value, report, started, passed=passed)

    value = copy.deepcopy(original)
    value["input"]["members"].reverse()
    _reject("timed-member-order", value, report, started, passed=passed)


def _observer_tests(passed: list[str]) -> None:
    report, started, original = _first("observer")
    value = copy.deepcopy(original)
    value["observer"]["zero_filled_gap_bytes"] += 1
    _reject("observer-gap-counter", value, report, started, observer=True, passed=passed)

    value = copy.deepcopy(original)
    events = value["observer"]["events"]
    if events:
        events.pop()
    else:
        value["observer"]["events_recorded"] = 1
    _reject("observer-event-count", value, report, started, observer=True, passed=passed)

    value = copy.deepcopy(original)
    value["observer"]["events_truncated"] = True
    _reject("observer-truncated-events", value, report, started, observer=True, passed=passed)


def _allocation_tests(passed: list[str]) -> None:
    report, started, original = _first("allocation-preflight")
    value = copy.deepcopy(original)
    value["metrics"]["counter_revision"] = "old-counter"
    _reject("allocation-counter-revision", value, report, started, allocation=True,
            passed=passed)

    value = copy.deepcopy(original)
    value["samples"][0]["allocation"]["status"] = "unavailable"
    _reject("allocation-unmeasured", value, report, started, allocation=True, passed=passed)

    value = copy.deepcopy(original)
    allocation = value["samples"][0]["allocation"]
    allocation["live_bytes_after"] += 1
    _reject("allocation-live-conservation", value, report, started, allocation=True,
            passed=passed)

    value = copy.deepcopy(original)
    allocation = value["samples"][0]["allocation"]
    allocation["region_peak_live_bytes"] = allocation["peak_live_bytes_after"] + 1
    _reject("allocation-region-peak-bound", value, report, started, allocation=True,
            passed=passed)


def _artifact_tests(passed: list[str]) -> None:
    candidates = sorted((RUNS / "qualification").glob("*/report.json"))
    require_case = next((path for path in candidates
                         if d.read(path.parent / "started.json")["case"].startswith("cfb-")),
                        None)
    if require_case is None:
        raise AssertionError("reader mutation tests require one CFB qualification report")
    report = require_case
    started = d.read(report.parent / "started.json")
    artifact_path = report.parent / "output.cfb"
    data = artifact_path.read_bytes()
    parsed = r.parse_cfb_artifact(data)
    assert parsed["bytes"] == len(data)
    assert parsed["sha256"] == r._sha256(data)
    assert parsed["header"]["sector_size"] in (512, 4096)
    assert parsed["directory"]
    if started["case"].startswith("cfb-"):
        identity, members = r._prepared_members(started["case"])
        actual = {row["name"]: row for row in parsed["streams"]}
        assert set(actual) == {name for name, _ in members}
        for name, payload in members:
            assert actual[name]["bytes"] == len(payload)
            assert actual[name]["sha256"] == r._sha256(payload)
        assert r._sequence_sha256((name, parsed["_stream_bytes"][name]) for name, _ in members) == identity["input_sha256"]
    damaged = bytearray(data)
    damaged[0] ^= 1
    try:
        r.parse_cfb_artifact(damaged)
    except r.ReaderError:
        passed.append("cfb-header-signature")
    else:
        raise AssertionError("CFB parser accepted a damaged signature")


def _statistical_tests(passed: list[str]) -> None:
    assert r.nearest_rank([4, 1, 3, 2], 0.5) == 2
    result = r.bootstrap_ratio([2] * 6, [1] * 6)
    assert result["estimate"] == 0.5
    assert result["ci95_low"] == result["ci95_high"] == 0.5
    interval = result["bootstrap"]
    assert interval["seed"] == 840084 and interval["resamples"] == 10_000
    assert interval["endpoint_indexes"] == [249, 9749]
    passed.append("bootstrap-seed-endpoints")


def main() -> None:
    passed: list[str] = []
    _timed_tests(passed)
    _observer_tests(passed)
    _allocation_tests(passed)
    _artifact_tests(passed)
    _statistical_tests(passed)
    d.write(P / "reader-tests.json", {
        "status": "pass",
        "mutations": passed,
        "statistical_checks": 2,
        "artifact_parser_checks": 4,
        "reader": d.desc(P / "readers.py"),
    })
    print("reader tests PASS", len(passed), "mutations")


if __name__ == "__main__":
    main()
