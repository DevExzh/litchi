#!/usr/bin/env python3
"""Mutation checks for the change-0458 report oracle.

The root smoke runner produces the four 30-sample reports consumed here.  A
successful run leaves only this receipt; each mutated report is confined to a
temporary directory and is removed before the receipt is written.
"""

from __future__ import annotations

import copy
import datetime
import hashlib
import importlib.util
import json
from pathlib import Path
import tempfile
from typing import Any, Callable


ROOT = Path(__file__).resolve().parent
RECEIPT = ROOT / "oracle-tests.json"
ORACLE_PATH = ROOT / "oracle.py"
SMOKE_PROTOCOL = {"samples": 30, "warmup": 3}
LANES = (
    ("normal-lifecycle", "normal", "lifecycle"),
    ("normal-phases", "normal", "phases"),
    ("allocator-lifecycle", "allocator", "lifecycle"),
    ("allocator-phases", "allocator", "phases"),
)


def _now() -> str:
    return datetime.datetime.now(datetime.timezone.utc).isoformat()


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def _load_oracle():
    spec = importlib.util.spec_from_file_location("oracle458_mutation_tests", ORACLE_PATH)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"cannot load {ORACLE_PATH}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def _lane(lane_id: str, instrumentation: str, scope: str) -> dict[str, str]:
    return {
        "id": lane_id,
        "repeat": "R1",
        "instrumentation": instrumentation,
        "shape": "tiny",
        "scope": scope,
    }


def _report_path(lane_id: str) -> Path:
    return ROOT / "smoke-r2" / f"{lane_id}.json"


def _read_report(path: Path) -> dict[str, Any]:
    with path.open() as stream:
        value = json.load(stream)
    if not isinstance(value, dict):
        raise AssertionError(f"{path} must contain a JSON object")
    return value


class TestFailure(AssertionError):
    pass


def _positive(oracle: Any, path: Path, lane: dict[str, str]) -> None:
    proof = oracle.validate(path, lane, SMOKE_PROTOCOL)
    if proof.get("status") != "pass":
        raise TestFailure(f"positive validation returned {proof!r}")
    expected_rows = SMOKE_PROTOCOL["samples"]
    if proof.get("rows") != expected_rows:
        raise TestFailure(f"positive validation returned wrong row count: {proof!r}")
    expected_phase_count = 5 if lane["scope"] == "phases" else 0
    if proof.get("phase_count") != expected_phase_count:
        raise TestFailure(f"positive validation returned wrong phase count: {proof!r}")


def _expect_rejected(
    oracle: Any,
    source_path: Path,
    lane: dict[str, str],
    temporary_path: Path,
    name: str,
    mutate: Callable[[dict[str, Any]], None],
    expected_fragment: str,
) -> None:
    value = copy.deepcopy(_read_report(source_path))
    mutate(value)
    temporary_path.write_text(json.dumps(value, indent=2) + "\n")
    try:
        oracle.validate(temporary_path, lane, SMOKE_PROTOCOL)
    except oracle.OracleError as error:
        message = str(error)
        if expected_fragment not in message:
            raise TestFailure(
                f"{name} rejected at an unexpected path: expected {expected_fragment!r}, "
                f"got {message!r}"
            ) from error
    else:
        raise TestFailure(f"{name} was accepted by the oracle")


def _output_bucket(length: int) -> str:
    if length == 0:
        return "bytes_0"
    if length <= 512:
        return "bytes_1_to_512"
    if length <= 4096:
        return "bytes_513_to_4096"
    if length <= 16384:
        return "bytes_4097_to_16384"
    if length <= 65536:
        return "bytes_16385_to_65536"
    return "bytes_over_65536"


def _run_case(
    tests: list[dict[str, Any]],
    diagnostics: list[str],
    name: str,
    function: Callable[[], None],
) -> None:
    try:
        function()
    except Exception as error:  # Keep running so the receipt records all failures.
        tests.append({"name": name, "status": "failed"})
        diagnostics.append(f"{name}: {type(error).__name__}: {error}")
    else:
        tests.append({"name": name, "status": "pass"})


def _write_receipt(value: dict[str, Any]) -> None:
    if RECEIPT.exists():
        raise FileExistsError(f"refusing to overwrite existing receipt: {RECEIPT}")
    try:
        with RECEIPT.open("x") as stream:
            json.dump(value, stream, indent=2, sort_keys=True)
            stream.write("\n")
    except FileExistsError:
        raise FileExistsError(f"refusing to overwrite existing receipt: {RECEIPT}")


def main() -> int:
    if RECEIPT.exists():
        raise SystemExit(f"refusing to overwrite existing receipt: {RECEIPT}")
    if not ORACLE_PATH.is_file():
        raise SystemExit(f"missing oracle: {ORACLE_PATH}")

    missing = [_report_path(lane_id) for lane_id, _, _ in LANES if not _report_path(lane_id).is_file()]
    if missing:
        paths = ", ".join(str(path.relative_to(ROOT)) for path in missing)
        raise SystemExit(f"smoke reports are not ready; missing {paths}")

    oracle = _load_oracle()
    tests: list[dict[str, Any]] = []
    diagnostics: list[str] = []
    started = _now()
    temporary_root: Path | None = None
    with tempfile.TemporaryDirectory(prefix="oracle-tests-", dir=ROOT) as temporary_directory:
        temporary_root = Path(temporary_directory)
        lane_values = {
            lane_id: _lane(lane_id, instrumentation, scope)
            for lane_id, instrumentation, scope in LANES
        }

        for lane_id, _, _ in LANES:
            source = _report_path(lane_id)
            lane = lane_values[lane_id]
            _run_case(
                tests,
                diagnostics,
                f"positive:{lane_id}",
                lambda source=source, lane=lane: _positive(oracle, source, lane),
            )

        normal_lifecycle = _report_path("normal-lifecycle")
        normal_lifecycle_lane = lane_values["normal-lifecycle"]
        normal_phases = _report_path("normal-phases")
        normal_phases_lane = lane_values["normal-phases"]
        allocator_lifecycle = _report_path("allocator-lifecycle")
        allocator_lifecycle_lane = lane_values["allocator-lifecycle"]
        allocator_phases = _report_path("allocator-phases")
        allocator_phases_lane = lane_values["allocator-phases"]

        def add(name: str, source: Path, lane: dict[str, str], mutate: Callable[[dict[str, Any]], None], fragment: str) -> None:
            _run_case(
                tests,
                diagnostics,
                f"mutation:{name}",
                lambda: _expect_rejected(
                    oracle,
                    source,
                    lane,
                    temporary_root / f"{name}.json",
                    name,
                    mutate,
                    fragment,
                ),
            )

        add(
            "changed-source-hash",
            normal_lifecycle,
            normal_lifecycle_lane,
            lambda value: value["rows"][0].__setitem__("source_sha256", "0" * 64),
            "report.rows[0].source_sha256",
        )
        add(
            "changed-output-hash",
            normal_lifecycle,
            normal_lifecycle_lane,
            lambda value: value["rows"][0].__setitem__("candidate_sha256", "0" * 64),
            "report.rows[0].candidate_sha256",
        )
        add(
            "false-preflight-gate",
            normal_lifecycle,
            normal_lifecycle_lane,
            lambda value: value["preflight_gates"].__setitem__("exact_noop_verified", False),
            "report.preflight_gates.exact_noop_verified",
        )
        add(
            "false-runtime-gate",
            normal_lifecycle,
            normal_lifecycle_lane,
            lambda value: value["rows"][0]["runtime_gates"].__setitem__("sink_digest_verified", False),
            "report.rows[0].runtime_gates.sink_digest_verified",
        )
        add(
            "boolean-sample-count",
            normal_lifecycle,
            normal_lifecycle_lane,
            lambda value: value.__setitem__("samples", True),
            "report.samples",
        )
        add(
            "wrong-sample-index",
            normal_lifecycle,
            normal_lifecycle_lane,
            lambda value: value["rows"][0].__setitem__("sample_index", 1),
            "report.rows[0].sample_index",
        )
        add(
            "swapped-phase-order",
            normal_phases,
            normal_phases_lane,
            lambda value: value["phase_order"].__setitem__(
                slice(0, 2), [value["phase_order"][1], value["phase_order"][0]]
            ),
            "report.phase_order",
        )
        add(
            "changed-phase-sum",
            normal_phases,
            normal_phases_lane,
            lambda value: value["rows"][0].__setitem__(
                "phase_sum_ns", value["rows"][0]["phase_sum_ns"] + 1
            ),
            "report.rows[0].phase_sum_ns",
        )

        normal_value = _read_report(normal_lifecycle)
        selected_bucket = _output_bucket(normal_value["identity"]["expected_output_bytes"])
        add(
            "bad-sink-bucket",
            normal_lifecycle,
            normal_lifecycle_lane,
            lambda value: value["rows"][0]["sink"]["write_size_buckets"].__setitem__(
                selected_bucket,
                value["rows"][0]["sink"]["write_size_buckets"][selected_bucket] + 1,
            ),
            f"report.rows[0].sink.write_size_buckets.{selected_bucket}",
        )
        add(
            "bad-sink-length",
            normal_lifecycle,
            normal_lifecycle_lane,
            lambda value: value["rows"][0]["sink"].__setitem__("accepted_bytes", 0),
            "report.rows[0].sink.accepted_bytes",
        )

        def bad_allocation_conservation(value: dict[str, Any]) -> None:
            metric = value["rows"][0]["lifecycle_allocation_metrics"]
            metric["allocated_bytes"] += 1

        def bad_allocation_peak(value: dict[str, Any]) -> None:
            metric = value["rows"][0]["lifecycle_allocation_metrics"]
            metric["region_peak_live_bytes"] = metric["peak_live_bytes_after"] + 1

        add(
            "bad-allocation-conservation",
            allocator_lifecycle,
            allocator_lifecycle_lane,
            bad_allocation_conservation,
            "report.rows[0].lifecycle_allocation_metrics",
        )
        add(
            "bad-allocation-peak",
            allocator_lifecycle,
            allocator_lifecycle_lane,
            bad_allocation_peak,
            "report.rows[0].lifecycle_allocation_metrics",
        )
        add(
            "wrong-instrumentation",
            allocator_lifecycle,
            allocator_lifecycle_lane,
            lambda value: value["allocator"].__setitem__("instrumentation", "none"),
            "report.allocator.instrumentation",
        )

        # Keep the smoke phase allocator lane in the matrix as a positive
        # control for allocator phase metrics; its mutation coverage is shared
        # with the direct lifecycle allocator counters above.
        _run_case(
            tests,
            diagnostics,
            "allocator-phase-metrics-present",
            lambda: _positive(oracle, allocator_phases, allocator_phases_lane),
        )

    # TemporaryDirectory removes all mutated reports before the receipt is
    # retained.  Keep this check outside the context manager.
    temporary_removed = temporary_root is not None and not temporary_root.exists()
    passed = sum(test["status"] == "pass" for test in tests)
    failed = len(tests) - passed
    receipt = {
        "schema": "litchi-0458-oracle-tests-v1",
        "change": 458,
        "status": "pass" if failed == 0 and not diagnostics and temporary_removed else "failed",
        "tests": tests,
        "diagnostics": diagnostics,
        "test_count": len(tests),
        "passed_tests": passed,
        "failed_tests": failed,
        "temporary_directory_removed": temporary_removed,
        "oracle_sha256": _sha256(ORACLE_PATH),
        "started_utc": started,
        "finished_utc": _now(),
    }
    _write_receipt(receipt)
    print(json.dumps(receipt, indent=2, sort_keys=True))
    return 0 if receipt["status"] == "pass" else 1


if __name__ == "__main__":
    raise SystemExit(main())
