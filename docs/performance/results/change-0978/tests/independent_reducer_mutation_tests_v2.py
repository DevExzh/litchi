#!/usr/bin/env python3
"""Independent offline mutation controls for the 0978 reducer contract.

The fixtures below are deliberately small and are built from the published
sidecar contracts. They test conservation, identity, owner proof, and plan
custody at the reducer boundary without replaying native capture or copying
the reducer's internal calculations. The script exits nonzero when an
invalid mutation is accepted; this is useful while the reducer is being
closed over the remaining custody gaps.
"""
from __future__ import annotations

import hashlib
import importlib.util
import json
import tempfile
from pathlib import Path
from typing import Any, Callable


REDUCER_PATH = Path("/home/zhuhe/litchi-performance-0978/tools/reduce_0978.py")
spec = importlib.util.spec_from_file_location("reduce_0978_independent_under_test", REDUCER_PATH)
if spec is None or spec.loader is None:
    raise RuntimeError(f"cannot load reducer at {REDUCER_PATH}")
reducer = importlib.util.module_from_spec(spec)
spec.loader.exec_module(reducer)


PERIOD = 1_000_000
MARKER = "doc_retention_probe_0942_cfb_profile::profiled_measured_case"
MARKER_DSO = "/sealed/target"


def _digest(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def _binding(path: Path) -> dict[str, Any]:
    return {"path": str(path.absolute()), "bytes": path.stat().st_size, "sha256": _digest(path)}


def _row(index: int, *, marker: bool = False, owner: str | None = None, phase: str = "startup", eligible: Any = False) -> dict[str, Any]:
    frames = []
    if marker:
        frames = [{"dso": MARKER_DSO, "symbol_base": MARKER, "resolved": owner is not None}]
    return {
        "index": index,
        "period": PERIOD,
        "time_ns": index,
        "pid": 7000,
        "tid": 7000,
        "cpu": 31,
        "frames": frames,
        "marker": marker,
        "physical_owner": owner,
        "inclusive_owners": [owner] if owner is not None else [],
        "phase": phase,
        "target_source_eligible": eligible,
    }


def _valid_actual() -> dict[str, Any]:
    owner = f"{MARKER_DSO}::{MARKER}"
    return {
        "D": 3 * PERIOD,
        "M": PERIOD,
        "R": 2 * PERIOD,
        "T": PERIOD,
        "Tknown": PERIOD,
        "U_target": 0,
        "U_marker": 0,
        "sample_count": 3,
        "marker_sample_count": 1,
        "physical_leaf_owner_period": {owner: PERIOD},
        "marker_known_owner_period": {owner: PERIOD},
        "physical_leaf_unknown_period": 0,
        "physical_leaf_unknown_samples": 0,
        "inclusive_owner_period": {owner: PERIOD},
        "target_known_owner_period": {owner: PERIOD},
        "target_source_owner_period": {owner: PERIOD},
        "phase_periods": {"target_execution": PERIOD},
        "phase_counts": {"target_execution": 1},
    }


def _valid_children() -> list[dict[str, Any]]:
    return [
        {
            "packet_id": f"child-{index:02d}",
            "child_number": index,
            "actual": _valid_actual(),
            "timing": [float(index), float(index + 1)],
            "artifacts": {},
        }
        for index in range(1, 13)
    ]


def _expect_rejection(label: str, operation: Callable[[], Any]) -> dict[str, Any]:
    try:
        operation()
    except reducer.ReducerFailure as error:
        return {"label": label, "status": "rejected", "code": error.code}
    except Exception as error:
        return {"label": label, "status": "unexpected_exception", "exception": type(error).__name__}
    return {"label": label, "status": "accepted_invalid"}


def _minimal_plan(root: Path, *, review_name: str = "wrong_review") -> Path:
    token = root / "bound.txt"
    token.write_text("bound\n", encoding="utf-8")
    bound = _binding(token)
    perf = Path("/usr/bin/perf")
    plan_path = root / "plan.json"
    plan = {
        "schema": "litchi.performance.0978.plan.mutation",
        "plan_path": str(plan_path.absolute()),
        "helper_seal": bound,
        "helper_seal_sha256": bound["sha256"],
        "perf_tool_drift": bound,
        "perf_requalification": bound,
        "perf_requalification_posthoc": bound,
        "perf": {
            "path": str(perf),
            "sha256": _digest(perf),
            "event": "instructions:u",
            "period": PERIOD,
            "callgraph": "dwarf,32768",
            "stack_cap": 32768,
        },
        "report_contract": {"case": "docfloat", "edits": 1, "warmups": 5, "samples_requested": 80},
        "child_gate": {
            "all_children": True,
            "raw_lost_sum": 0,
            "terminal_lost_samples": 0,
            "min_marker_samples": 100,
            "on_failure": "stop_without_retry",
        },
        "batch": {"children": 12, "serial": True, "period": PERIOD, "retry": False},
        "decoder_reviews": {review_name: bound},
    }
    plan_path.write_text(json.dumps(plan, indent=2) + "\n", encoding="utf-8")
    return plan_path


def _minimal_artifact_map(root: Path, *, include_custody: bool = False) -> dict[str, Any]:
    artifacts: dict[str, Any] = {}
    for name in ("corrected_script", "corrected_log", "perf_mmap_dump", "perf_evlist", "raw_stack", "lifecycle_ledger", "attribution_samples"):
        path = root / f"{name}.dat"
        path.write_text(name + "\n", encoding="utf-8")
        artifacts[name] = _binding(path)
    if include_custody:
        for name in ("abi_scan", "report", "loss", "terminal_summary", "terminal", "launch", "command", "perf_data", "perf_stderr", "bridge", "affinity_observations"):
            path = root / f"{name}.dat"
            path.write_text(name + "\n", encoding="utf-8")
            artifacts[name] = _binding(path)
    return {"packet_id": "child-01", "artifacts": artifacts}


def _raw_reparse_mutation(*, mismatch: bool = True) -> None:
    with tempfile.TemporaryDirectory(prefix="litchi-0978-reducer-raw-") as temp:
        root = Path(temp)
        dump = root / "dump.txt"
        evlist = root / "evlist.txt"
        raw_sidecar = root / "raw.json"
        dump.write_text("raw\n", encoding="utf-8")
        evlist.write_text("evlist\n", encoding="utf-8")
        raw_sidecar.write_text("{}\n", encoding="utf-8")
        contract = {"period": PERIOD, "stack_cap": 32768}
        sample = {"index": 1, "period": PERIOD, "time_ns": 1, "pid": 7, "tid": 7, "cpu": 31, "frames": []}

        class FakeScanner:
            @staticmethod
            def parse_evlist_contract(*_args: Any, **_kwargs: Any) -> dict[str, Any]:
                return contract

            @staticmethod
            def scan_dump(*_args: Any, **_kwargs: Any) -> dict[str, Any]:
                return {
                    "status": "complete",
                    "sample_conservation": True,
                    "period_valid": True,
                    "pid_tid_cpu_valid": True,
                    "time_valid": True,
                    "bounds_valid": True,
                    "physical_identity_valid": True,
                    "violations": [],
                    "sample_count": 1,
                    "sample_ledger": [sample],
                    # A real scanner found an unrelated raw record. That is
                    # valid, but the sidecar's full census must agree.
                    "raw_event_count": 2,
                    "physical_location_count": 2,
                    "raw_sample_record_count": 1,
                    "raw_unconsumed_count": 1,
                    "sample_line_count": 1,
                    "raw_record_type_counts": {"1": 1, "9": 1},
                }

        original = reducer._load_sibling_module
        reducer._load_sibling_module = lambda *_args, **_kwargs: FakeScanner
        try:
            artifacts = {"perf_mmap_dump": _binding(dump), "perf_evlist": _binding(evlist)}
            raw_obj = {
                "event_contract": contract,
                "dump": _binding(dump),
                "expected_pid": 7,
                "expected_cpu": 31,
                "sample_count": 1,
                "total_period": PERIOD,
                "raw_event_count": 1 if mismatch else 2,
                "physical_location_count": 1 if mismatch else 2,
                "raw_sample_record_count": 1,
                "raw_unconsumed_count": 0 if mismatch else 1,
                "sample_line_count": 1,
                "raw_record_type_counts": {"9": 1} if mismatch else {"1": 1, "9": 1},
            }
            reducer._reparse_raw_scanner(raw_sidecar, raw_obj, artifacts, marker=MARKER, marker_dso=MARKER_DSO, allow_synthetic=False)
        finally:
            reducer._load_sibling_module = original


def _lifecycle_identity_mutation(*, mismatch: bool = True) -> None:
    with tempfile.TemporaryDirectory(prefix="litchi-0978-reducer-life-") as temp:
        dump = Path(temp) / "dump.txt"
        dump.write_text("raw\n", encoding="utf-8")
        lifecycle_row = {
            "index": 1, "period": PERIOD, "time_ns": 1, "pid": 7, "tid": 7, "cpu": 31,
            "phase": "startup", "mode": "user", "marker": False,
            "physical_owner": None, "target_source_eligible": False,
            "unknown": False, "unknown_reason": None,
        }
        parser_sample = dict(lifecycle_row)
        if mismatch:
            parser_sample.update({"index": 99, "period": PERIOD + 1, "time_ns": 999})
        built_source = parser_sample if mismatch else lifecycle_row
        built_row = {
            "raw_identity": dict(built_source),
            "phase": built_source["phase"], "mode": built_source["mode"],
            "marker": built_source["marker"], "physical_owner": built_source["physical_owner"],
            "target_source_eligible": built_source["target_source_eligible"],
            "unknown": built_source["unknown"], "unknown_reason": built_source["unknown_reason"],
        }

        class FakeLifecycle:
            @staticmethod
            def parse_lifecycle_dump(*_args: Any, **_kwargs: Any) -> dict[str, Any]:
                return {
                    "status": "complete",
                    "raw_metadata_crosscheck": True,
                    "unique_record_offsets": True,
                    "raw_lost_sum": 0,
                    "terminal_lost_values": [0],
                    "sample_descriptor_count": 1,
                    "sample_records": [parser_sample],
                }

            @staticmethod
            def build_ledger(*_args: Any, **_kwargs: Any) -> dict[str, Any]:
                return {"status": "complete", "samples": [built_row], "ledger": {}}

        original = reducer._load_sibling_module
        reducer._load_sibling_module = lambda *_args, **_kwargs: FakeLifecycle
        try:
            reducer._reparse_lifecycle_stream(
                dump,
                {"expected_pid": 7, "expected_cpu": 31},
                {"target": {"path": "/sealed/target", "build_id": "abcdef12", "comm": "target"}, "terminal_lost_samples": 0},
                [lifecycle_row],
                raw_scanner_rows=[lifecycle_row],
                attribution_obj={"attributions": {"1": {}}},
                packet={"gate": {"terminal_lost_samples": 0}},
                plan={"binary": {"path": "/sealed/target"}, "marker": {"build_id": "abcdef12"}, "workload": {"target_comm": "target"}},
                marker=MARKER,
                marker_dso=MARKER_DSO,
                allow_synthetic=False,
            )
        finally:
            reducer._load_sibling_module = original


def main() -> int:
    baseline = reducer.reduce_children(_valid_children())
    assert baseline["status"] == "complete"
    assert baseline["summary"]["period_conservation"]["D_equals_M_plus_R"] is True

    # Positive controls make the negative cases meaningful and catch fixture
    # mistakes independently of the reducer's mutation outcomes.
    owner = f"{MARKER_DSO}::{MARKER}"
    valid_rows = [_row(1), _row(2, marker=True, owner=owner, phase="target_execution", eligible=True), _row(3)]
    valid_ledger = {
        "D": 3 * PERIOD, "M": PERIOD, "R": 2 * PERIOD, "T": PERIOD,
        "U_target": 0, "U_marker": 0,
        "target_known_owner_period": {owner: PERIOD},
        "target_source_eligible_period": {owner: PERIOD},
        "marker_known_owner_period": {owner: PERIOD},
    }
    derived = reducer._derive_child(valid_rows, valid_rows, valid_rows, valid_ledger, marker_dso=MARKER_DSO)
    assert derived["D"] == derived["M"] + derived["R"]
    reducer._normalise_sample(valid_rows[1], 2, "positive", marker=MARKER, marker_dso=MARKER_DSO)
    _raw_reparse_mutation(mismatch=False)
    _lifecycle_identity_mutation(mismatch=False)
    reducer._compare_identities(
        [{"index": 1, "period": PERIOD, "time_ns": 1, "pid": 7, "tid": 7, "cpu": 31}],
        [{"index": 1, "period": PERIOD, "time_ns": 1, "pid": 7, "tid": 7, "cpu": 31}],
        "raw", "lifecycle",
    )
    with tempfile.TemporaryDirectory(prefix="litchi-0978-reducer-artifact-control-") as temp:
        reducer._artifact_map(_minimal_artifact_map(Path(temp), include_custody=True))

    mutations: list[dict[str, Any]] = []

    dmr_children = _valid_children()
    dmr_children[0]["actual"]["R"] = 0
    mutations.append(_expect_rejection("per-child D/M/R conservation", lambda: reducer.reduce_children(dmr_children)))

    rows = [{
        "index": 1, "period": PERIOD, "time_ns": 1, "pid": 7, "tid": 7, "cpu": 31,
        "phase": "startup", "mode": "user", "marker": False,
        "target_source_eligible": False, "unknown": False,
    }]
    mutations.append(_expect_rejection(
        "required lifecycle D/M/R/T/unknown ledger fields",
        lambda: reducer._require_live_lifecycle_contract({"schema": "litchi.performance.0978.lifecycle_ledger.v1"}, rows, {}),
    ))

    summary_only = {"index": 1, "period": PERIOD, "time_ns": 1, "pid": 7, "tid": 7, "cpu": 31, "marker": True, "physical_owner": f"{MARKER_DSO}::{MARKER}", "inclusive_owners": [f"{MARKER_DSO}::{MARKER}"]}
    mutations.append(_expect_rejection(
        "attribution owner proof requires full frames",
        lambda: reducer._normalise_sample(summary_only, 1, "attribution", marker=MARKER, marker_dso=MARKER_DSO, allow_summary_claims=True),
    ))

    target_without_eligibility = _row(1, marker=False, owner=f"{MARKER_DSO}::target", phase="target_execution", eligible=None)
    target_without_eligibility["physical_owner"] = f"{MARKER_DSO}::target"
    target_without_eligibility["frames"] = [{"dso": MARKER_DSO, "symbol_base": "target", "resolved": True}]
    target_ledger = {
        "D": PERIOD, "M": 0, "R": PERIOD, "T": PERIOD,
        "U_target": PERIOD, "U_marker": 0, "marker_sample_count": 0,
        "target_known_owner_period": {}, "target_source_eligible_period": {},
        "marker_known_owner_period": {},
        "conservation": {"D_equals_M_plus_R": True, "T_equals_known_plus_unknown": True},
    }
    mutations.append(_expect_rejection(
        "target source eligibility is explicit",
        lambda: reducer._require_live_lifecycle_contract({"schema": "litchi.performance.0978.lifecycle_ledger.v1"}, [target_without_eligibility], target_ledger),
    ))

    mutations.append(_expect_rejection("raw reparse compares full raw census", _raw_reparse_mutation))
    mutations.append(_expect_rejection("lifecycle reparse compares sample identities", _lifecycle_identity_mutation))

    with tempfile.TemporaryDirectory(prefix="litchi-0978-reducer-plan-") as temp:
        root = Path(temp)
        plan_path = root / "plan.json"
        plan_path.write_text("{}\n", encoding="utf-8")
        plan = {"schema": "litchi.performance.0978.plan.v1"}
        valid_packet = {"packet_id": "child-01", "plan": {"path": str(plan_path.absolute()), "sha256": _digest(plan_path), "schema": plan["schema"]}}
        reducer._check_packet_plan_binding(valid_packet, plan_path, plan)
        mutations.append(_expect_rejection("packet plan binding is required", lambda: reducer._check_packet_plan_binding({"packet_id": "child-01"}, plan_path, plan)))

    print(json.dumps({
        "schema": "litchi.performance.0978.independent_reducer_mutation_tests.v2",
        "source_sha256": _digest(REDUCER_PATH),
        "baseline": "PASS",
        "mutations": mutations,
        "accepted_invalid_count": sum(item["status"] == "accepted_invalid" for item in mutations),
        "unexpected_exception_count": sum(item["status"] == "unexpected_exception" for item in mutations),
    }, indent=2, sort_keys=True))
    return 1 if any(item["status"] != "rejected" for item in mutations) else 0


if __name__ == "__main__":
    raise SystemExit(main())
