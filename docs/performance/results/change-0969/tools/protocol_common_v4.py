#!/usr/bin/env python3
"""Shared fail-closed helpers for the prospective 0969 CFB protocol.

This module only validates external JSON custody and reduces retained packets.
It never invokes Cargo or a benchmark executable.
"""
from __future__ import annotations

import hashlib
import json
import math
import random
from pathlib import Path
from typing import Any, Callable

RUNNER_SCHEMA = "litchi.performance.0969.runner.v4"
PROBE_SCHEMA = "litchi.performance.0917.doc_retention.v1"
PLAN_SCHEMA = "litchi.performance.0969.plan.v4"
DECISION_SCHEMA = "litchi.performance.0969.perf_decision.v4"
AA_SCHEMA = "litchi.performance.0969.aa_readiness.v4"
GO_SCHEMA = "litchi.performance.0969.root_go.v1"
HELPER_NAMES = ("protocol_common", "runner", "reducer", "plan_generator")
BOOTSTRAP_SEED = 969004
BOOTSTRAP_RESAMPLES = 10_000
CI_LIMIT = 1.05
NATIVE_PAIRS = 12
ALLOCATOR_PAIRS = 3
NATIVE_WARMUPS = 5
NATIVE_SAMPLES = 80
ALLOCATOR_WARMUPS = 0
ALLOCATOR_SAMPLES = 1
CASES = ("docnohf", "docfloat")
EDITS = (1, 2, 8, 32)
RETENTIONS = ("zero", "default", "release")
ALLOC_FIELDS = (
    "allocated_bytes",
    "deallocated_bytes",
    "allocation_calls",
    "peak_live_bytes",
    "retained_bytes",
)
CORRECTNESS_FIELDS = (
    "all_stages_change_text",
    "exact_cross_route_outputs",
    "final_paragraph_semantic_readback",
    "final_public_reader_reopen",
    "exact_noop_source",
    "noop_shares_source_allocation",
    "forward_patch_exact_output",
    "inverse_patch_exact_source",
    "foreign_patch_refused",
    "stale_patch_refused",
    "same_value_preserves_retention",
    "operation_limit_refusal_atomic",
    "retained_capacities_within_ceiling",
    "zero_route_clears_retention",
    "release_route_clears_every_stage",
    "untouched_streams_preserved",
    "directory_metadata_preserved",
    "semantic_oracles_passed",
    "semantic_change_count_matches_edit_count",
    "allocator_retained_matches_output_capacity",
)
ORACLE_BOOL_FIELDS = (
    "oracle_ok",
    "source_structurally_valid",
    "expected_structurally_valid",
    "output_structurally_valid",
    "semantic_reopen_ok",
    "semantic_directory_metadata_matches_expected",
    "unchanged_stream_bytes_match_expected",
    "unchanged_source_output_stream_bytes_match",
    "output_stream_paths_match_expected",
    "output_stream_bytes_match_expected",
    "root_clsid_preserved",
    "storage_clsids_preserved",
)
ORACLE_FIELDS = set(ORACLE_BOOL_FIELDS) | {"failure_reasons"}


class ProtocolError(ValueError):
    """An input cannot be admitted to the protocol."""


def load_json(path: Path) -> dict[str, Any]:
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise ProtocolError(f"{path}: invalid JSON: {exc}") from exc
    if not isinstance(value, dict):
        raise ProtocolError(f"{path}: expected a JSON object")
    return value


def write_json_exclusive(path: Path, value: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    try:
        with path.open("x", encoding="utf-8") as handle:
            json.dump(value, handle, indent=2, sort_keys=True)
            handle.write("\n")
    except FileExistsError as exc:
        raise ProtocolError(f"refusing to overwrite existing artifact: {path}") from exc


def sha256(path: Path) -> str:
    if not path.is_file():
        raise ProtocolError(f"missing file: {path}")
    digest = hashlib.sha256()
    try:
        with path.open("rb") as handle:
            for chunk in iter(lambda: handle.read(1024 * 1024), b""):
                digest.update(chunk)
    except OSError as exc:
        raise ProtocolError(f"cannot read {path}: {exc}") from exc
    return digest.hexdigest()


def check_file(path_value: Any, expected_sha: Any, label: str, *, allow_unbound: bool) -> None:
    if not isinstance(path_value, str) or not path_value:
        if allow_unbound:
            return
        raise ProtocolError(f"{label}: path is not bound")
    path = Path(path_value)
    if not path.is_file():
        if allow_unbound:
            return
        raise ProtocolError(f"{label}: missing {path}")
    if not isinstance(expected_sha, str) or len(expected_sha) != 64:
        if allow_unbound:
            return
        raise ProtocolError(f"{label}: SHA-256 is not bound")
    actual = sha256(path)
    if actual != expected_sha:
        raise ProtocolError(f"{label}: SHA-256 mismatch ({actual} != {expected_sha})")


def check_pass_receipt(path_value: Any, expected_sha: Any, label: str, *, allow_unbound: bool) -> dict[str, Any] | None:
    """Bind a review/receipt file and require its emitted status to pass."""
    check_file(path_value, expected_sha, label, allow_unbound=allow_unbound)
    if not isinstance(path_value, str) or not path_value or not Path(path_value).is_file():
        return None
    receipt = load_json(Path(path_value))
    status = receipt.get("status")
    if not isinstance(status, str) or not status.startswith("pass"):
        raise ProtocolError(f"{label}: bound receipt status is not pass ({status!r})")
    return receipt


def _same_binding(actual: Any, expected: dict[str, Any], label: str) -> None:
    if not isinstance(actual, dict):
        raise ProtocolError(f"GO {label}: binding section is missing")
    for key, row in expected.items():
        observed = actual.get(key)
        if not isinstance(observed, dict):
            raise ProtocolError(f"GO {label}/{key}: binding is missing")
        for field, value in row.items():
            if observed.get(field) != value:
                raise ProtocolError(f"GO {label}/{key}/{field}: binding mismatch")


def validate_go(go_path: Path, plan_path: Path, plan: dict[str, Any], *, binary_roles: tuple[str, ...]) -> dict[str, Any]:
    """Validate an explicit root GO against the immutable plan custody."""
    go_path = go_path.resolve()
    go = load_json(go_path)
    if go.get("schema") != GO_SCHEMA or go.get("authorized") is not True:
        raise ProtocolError("root GO schema/authorization is not valid")
    if (go.get("quiescence") or {}).get("status") != "pass":
        raise ProtocolError("root GO lacks a passing no-build/no-workload quiescence admission")
    plan_path = plan_path.resolve()
    plan_bytes = plan_path.stat().st_size
    plan_binding = go.get("plan") or {}
    if plan_binding.get("path") != str(plan_path) or plan_binding.get("bytes") != plan_bytes or plan_binding.get("sha256") != sha256(plan_path):
        raise ProtocolError("root GO plan path/size/SHA binding mismatch")
    review_bindings = {
        name: {"path": plan[name]["path"], "sha256": plan[name]["sha256"]}
        for name in ("design_review", "baseline_reuse_review")
    }
    _same_binding(go.get("review_bindings"), review_bindings, "reviews")
    helpers = plan.get("helpers") or {}
    _same_binding(go.get("helper_bindings"), {name: {"path": helpers[name]["path"], "sha256": helpers[name]["sha256"]} for name in HELPER_NAMES}, "helpers")
    source = plan.get("source") or {}
    _same_binding(
        go.get("source_bindings"),
        {name: {"manifest": source[name]["manifest"], "manifest_sha256": source[name]["manifest_sha256"]} for name in ("baseline", "candidate")},
        "sources",
    )
    binaries = plan.get("binaries") or {}
    _same_binding(
        go.get("binary_bindings"),
        {name: {"path": binaries[name]["path"], "sha256": binaries[name]["sha256"]} for name in binary_roles},
        "binaries",
    )
    return {"path": str(go_path), "sha256": sha256(go_path), "receipt": go}


def percentile(values: list[float], quantile: float) -> float:
    if not values:
        raise ProtocolError("percentile of empty values")
    ordered = sorted(float(value) for value in values)
    position = (len(ordered) - 1) * quantile
    lower = math.floor(position)
    upper = math.ceil(position)
    if lower == upper:
        return ordered[lower]
    return ordered[lower] + (ordered[upper] - ordered[lower]) * (position - lower)


def bootstrap_median(values: list[float], seed: int = BOOTSTRAP_SEED, resamples: int = BOOTSTRAP_RESAMPLES) -> dict[str, float]:
    if not values:
        raise ProtocolError("bootstrap of empty values")
    if resamples <= 0:
        raise ProtocolError("bootstrap resample count must be positive")
    rng = random.Random(seed)
    medians = []
    for _ in range(resamples):
        sample = [values[rng.randrange(len(values))] for _ in values]
        medians.append(percentile(sample, 0.5))
    return {
        "median": percentile(values, 0.5),
        "lower": percentile(medians, 0.025),
        "upper": percentile(medians, 0.975),
    }


def ratio(numerator: float, denominator: float, label: str) -> float:
    if denominator <= 0:
        raise ProtocolError(f"{label}: baseline metric is not positive")
    value = numerator / denominator
    if not math.isfinite(value) or value <= 0:
        raise ProtocolError(f"{label}: invalid ratio {value!r}")
    return value


def expected_cells() -> list[str]:
    return [f"{case}/e{edits}/{retention}" for case in CASES for edits in EDITS for retention in RETENTIONS]


def cell_id(cell: dict[str, Any]) -> str:
    return f"{cell.get('case')}/e{cell.get('edits')}/{cell.get('retention')}"


def fixture_retention(value: str) -> str:
    mapping = {"zero": "zero", "default": "default_8mib", "release": "release_8mib"}
    try:
        return mapping[value]
    except KeyError as exc:
        raise ProtocolError(f"unknown retention {value!r}") from exc


def check_correctness(value: Any, label: str) -> None:
    if not isinstance(value, dict) or set(value) != set(CORRECTNESS_FIELDS):
        raise ProtocolError(f"{label}: expected exactly the sealed correctness gate set")
    if any(type(value[name]) is not bool or value[name] is not True for name in CORRECTNESS_FIELDS):
        raise ProtocolError(f"{label}: every sealed correctness gate must be boolean true")


def check_oracle(value: Any, label: str) -> None:
    if not isinstance(value, dict) or set(value) != ORACLE_FIELDS:
        raise ProtocolError(f"{label}: expected exactly the sealed oracle field set")
    if any(type(value[name]) is not bool or value[name] is not True for name in ORACLE_BOOL_FIELDS):
        raise ProtocolError(f"{label}: every oracle boolean must be true")
    if not isinstance(value["failure_reasons"], list) or value["failure_reasons"]:
        raise ProtocolError(f"{label}: failure_reasons must be an empty list")


def positive_int(value: Any, label: str) -> None:
    if type(value) is not int or value <= 0:
        raise ProtocolError(f"{label}: expected a positive integer")


def nonnegative_int(value: Any, label: str) -> None:
    if type(value) is not int or value < 0:
        raise ProtocolError(f"{label}: expected a non-negative integer")


def sha256_text(value: Any, label: str) -> None:
    if not isinstance(value, str) or len(value) != 64:
        raise ProtocolError(f"{label}: expected a SHA-256 string")
    try:
        int(value, 16)
    except ValueError as exc:
        raise ProtocolError(f"{label}: invalid SHA-256 string") from exc


def check_plan(
    plan_path: Path,
    *,
    allow_unbound: bool = False,
    for_capture: bool = False,
    for_aa: bool = False,
) -> dict[str, Any]:
    if for_capture and for_aa:
        raise ProtocolError("for_capture and for_aa are mutually exclusive")
    plan = load_json(plan_path)
    if plan.get("schema") != PLAN_SCHEMA:
        raise ProtocolError(f"plan schema must be {PLAN_SCHEMA}")
    protocol = plan.get("protocol_document") or {}
    check_file(protocol.get("path"), protocol.get("sha256"), "protocol document", allow_unbound=allow_unbound)
    if plan.get("status") == "capture_authorized" and plan.get("no_build_or_workload") is True:
        raise ProtocolError("capture_authorized plan still says no_build_or_workload")
    if plan.get("execution", {}).get("serial") is not True:
        raise ProtocolError("execution.serial must be true")
    if plan.get("execution", {}).get("retry") is not False:
        raise ProtocolError("execution.retry must be false")
    if plan.get("execution", {}).get("quiescence_required") is not True:
        raise ProtocolError("execution.quiescence_required must be true")

    for name in ("design_review", "baseline_reuse_review"):
        row = plan.get(name) or {}
        check_pass_receipt(row.get("path"), row.get("sha256"), name, allow_unbound=allow_unbound)
        if row.get("status") not in {"pass", "passed"}:
            raise ProtocolError(f"{name}: status must be pass")

    helpers = plan.get("helpers") or {}
    if set(helpers) != set(HELPER_NAMES):
        raise ProtocolError("helper binding set drifted")
    for name in HELPER_NAMES:
        row = helpers[name]
        check_file(row.get("path"), row.get("sha256"), f"helper {name}", allow_unbound=allow_unbound)

    source = plan.get("source") or {}
    for role in ("baseline", "candidate"):
        row = source.get(role) or {}
        check_file(row.get("manifest"), row.get("manifest_sha256"), f"source.{role}.manifest", allow_unbound=allow_unbound)
    fixtures = plan.get("fixtures") or {}
    if set(fixtures) != set(CASES):
        raise ProtocolError(f"fixture set must be {set(CASES)}")
    for case in CASES:
        row = fixtures[case]
        check_file(row.get("path"), row.get("sha256"), f"fixture {case}", allow_unbound=allow_unbound)
        path = Path(row.get("path", ""))
        if path.is_file() and row.get("bytes") != path.stat().st_size:
            raise ProtocolError(f"fixture {case}: byte count changed")

    matrix = plan.get("real_matrix") or {}
    cells = matrix.get("cells") or []
    ids = [cell_id(row) for row in cells]
    if len(cells) != 24 or set(ids) != set(expected_cells()) or len(set(ids)) != 24:
        raise ProtocolError("real matrix must be exactly the 24 declared cells")
    native = matrix.get("native") or {}
    if (native.get("pairs_per_cell"), native.get("warmups"), native.get("samples")) != (NATIVE_PAIRS, NATIVE_WARMUPS, NATIVE_SAMPLES):
        raise ProtocolError("native schedule drifted")
    allocator = matrix.get("allocator") or {}
    if (allocator.get("pairs_per_cell"), allocator.get("warmups"), allocator.get("samples")) != (ALLOCATOR_PAIRS, ALLOCATOR_WARMUPS, ALLOCATOR_SAMPLES):
        raise ProtocolError("allocator schedule drifted")
    for lane in (native, allocator):
        if lane.get("relative_ratio") != "candidate_over_baseline":
            raise ProtocolError("relative ratio must be candidate_over_baseline")
        if lane.get("bootstrap_seed") != BOOTSTRAP_SEED or lane.get("bootstrap_resamples") != BOOTSTRAP_RESAMPLES:
            raise ProtocolError("bootstrap seed/resample count drifted")
    if native.get("noninferiority_upper_ci") != CI_LIMIT:
        raise ProtocolError("native non-inferiority limit drifted")
    if allocator.get("requested_metric") != "samples[i].allocation.allocated_bytes":
        raise ProtocolError("requested metric must bind the existing allocated_bytes field")
    if allocator.get("rss_metric") != "packet.rss.kb":
        raise ProtocolError("RSS metric must bind packet.rss.kb")

    aa = plan.get("aa_noise") or {}
    if aa.get("cells") != expected_cells() or aa.get("coverage") != "all_real_cells":
        raise ProtocolError("A/A coverage must be all 24 real cells")
    if (aa.get("native_pairs_per_cell"), aa.get("warmups"), aa.get("samples")) != (NATIVE_PAIRS, NATIVE_WARMUPS, NATIVE_SAMPLES):
        raise ProtocolError("A/A schedule drifted")
    if aa.get("same_binary") != "baseline" or aa.get("order") != ["AB", "BA"]:
        raise ProtocolError("A/A binary/order drifted")
    if aa.get("bootstrap_seed") != BOOTSTRAP_SEED or aa.get("bootstrap_resamples") != BOOTSTRAP_RESAMPLES or aa.get("upper_ci_limit") != CI_LIMIT:
        raise ProtocolError("A/A threshold or bootstrap drifted")
    if aa.get("before_candidate_formal_launch") is not True:
        raise ProtocolError("A/A must gate candidate formal launch")
    benefit = plan.get("benefit") or {}
    if benefit.get("requested_metric") != "samples[i].allocation.allocated_bytes":
        raise ProtocolError("benefit requested metric drifted")
    if benefit.get("requested_endpoint_median_upper") != 0.95:
        raise ProtocolError("requested-byte benefit threshold drifted")
    if benefit.get("native_p50_median_upper") != 0.95 or benefit.get("native_p50_ci_upper") != 1.0:
        raise ProtocolError("native benefit thresholds drifted")

    binaries = plan.get("binaries") or {}
    binary_roles = ("native_baseline",) if for_aa else ("native_baseline", "native_candidate", "allocator_baseline", "allocator_candidate")
    for role in binary_roles:
        row = binaries.get(role) or {}
        check_file(row.get("path"), row.get("sha256"), f"binary {role}", allow_unbound=allow_unbound)
    if for_aa:
        if plan.get("status") != "aa_authorized":
            raise ProtocolError("A/A capture requires aa_authorized status")
        row = binaries.get("native_baseline") or {}
        check_file(row.get("path"), row.get("sha256"), "binary native_baseline", allow_unbound=False)
    if for_capture:
        if plan.get("status") != "capture_authorized":
            raise ProtocolError("formal capture requires capture_authorized status")
        readiness = plan.get("aa_readiness_receipt") or {}
        receipt = check_pass_receipt(
            readiness.get("path"), readiness.get("sha256"), "A/A readiness receipt", allow_unbound=False
        )
        if not isinstance(receipt, dict) or receipt.get("schema") != AA_SCHEMA or receipt.get("status") != "pass":
            raise ProtocolError("A/A readiness receipt is not a passing v4 receipt")
        if receipt.get("bootstrap_seed") != BOOTSTRAP_SEED or receipt.get("bootstrap_resamples") != BOOTSTRAP_RESAMPLES:
            raise ProtocolError("A/A readiness receipt bootstrap binding drifted")
        if set((receipt.get("cells") or {}).keys()) != set(expected_cells()):
            raise ProtocolError("A/A readiness receipt does not cover all 24 cells")
        aa_manifest = plan.get("aa_capture_manifest") or {}
        check_file(aa_manifest.get("path"), aa_manifest.get("sha256"), "A/A capture manifest", allow_unbound=False)
        manifest = load_json(Path(aa_manifest["path"]))
        if manifest.get("schema") != RUNNER_SCHEMA or manifest.get("no_retry") is not True:
            raise ProtocolError("A/A capture manifest schema/retry binding drifted")
        packet_paths = manifest.get("packets")
        if not isinstance(packet_paths, list) or len(packet_paths) != 2 * NATIVE_PAIRS * len(expected_cells()):
            raise ProtocolError("A/A capture manifest packet count drifted")
        if len(set(packet_paths)) != len(packet_paths):
            raise ProtocolError("A/A capture manifest contains duplicate packet paths")
        for role in ("native_baseline", "native_candidate", "allocator_baseline", "allocator_candidate"):
            row = binaries.get(role) or {}
            check_file(row.get("path"), row.get("sha256"), f"binary {role}", allow_unbound=False)
    return plan


def check_probe_report(report: dict[str, Any], *, cell: dict[str, Any], fixture: dict[str, Any], allocator: bool) -> None:
    if report.get("schema") != PROBE_SCHEMA:
        raise ProtocolError("probe report schema mismatch")
    case = cell.get("case")
    if report.get("case") != case or report.get("edits") != cell.get("edits"):
        raise ProtocolError("probe report cell mismatch")
    if report.get("retention") != fixture_retention(str(cell.get("retention"))):
        raise ProtocolError("probe report retention mismatch")
    if report.get("source_sha256") != fixture.get("sha256") or report.get("source_bytes") != fixture.get("bytes"):
        raise ProtocolError("probe fixture identity mismatch")
    sha256_text(report.get("source_sha256"), "probe.source_sha256")
    positive_int(report.get("source_bytes"), "probe.source_bytes")
    if type(report.get("edits")) is not int or report.get("edits") != cell.get("edits"):
        raise ProtocolError("probe.edits identity mismatch")
    warmups = ALLOCATOR_WARMUPS if allocator else NATIVE_WARMUPS
    samples = ALLOCATOR_SAMPLES if allocator else NATIVE_SAMPLES
    if type(report.get("warmups")) is not int or type(report.get("samples_requested")) is not int or report.get("warmups") != warmups or report.get("samples_requested") != samples:
        raise ProtocolError("probe warmup/sample count mismatch")
    if type(report.get("allocator_instrumented")) is not bool or type(report.get("timing_claim")) is not bool or report.get("allocator_instrumented") != allocator or report.get("timing_claim") == allocator:
        raise ProtocolError("probe lane flags mismatch")
    check_correctness(report.get("correctness"), "probe.correctness")
    for name in ("output_sha256", "reference_output_sha256"):
        sha256_text(report.get(name), f"probe.{name}")
    for name in ("output_bytes", "output_capacity", "reference_output_bytes"):
        nonnegative_int(report.get(name), f"probe.{name}")
    elapsed = report.get("elapsed_ns")
    if not isinstance(elapsed, list) or len(elapsed) != (0 if allocator else samples):
        raise ProtocolError("probe elapsed count mismatch")
    if any(type(value) is not int or value <= 0 for value in elapsed):
        raise ProtocolError("probe elapsed_ns contains a non-positive or non-integer value")
    rows = report.get("samples")
    if not isinstance(rows, list) or len(rows) != samples:
        raise ProtocolError("probe sample row count mismatch")
    routes = report.get("route_outputs")
    if not isinstance(routes, list) or len(routes) != 3:
        raise ProtocolError("probe route output set incomplete")
    for route in routes:
        if not isinstance(route, dict):
            raise ProtocolError("probe.route: expected object")
        check_oracle(route.get("oracle"), "probe.route.oracle")
    for index, row in enumerate(rows):
        # These are the retained 0917 harness field names.  Do not invent a
        # second reopen/output oracle for the 0969 adapter.
        if (
            not isinstance(row, dict)
            or type(row.get("index")) is not int
            or row.get("index") != index
            or row.get("exact_output_equal_reference") is not True
            or row.get("final_public_reader_reopen") is not True
            or row.get("final_paragraph_semantic_readback") is not True
        ):
            raise ProtocolError("probe output/reopen oracle failed")
        sha256_text(row.get("output_sha256"), "probe.sample.output_sha256")
        if row.get("output_sha256") != report.get("output_sha256"):
            raise ProtocolError("probe sample output identity mismatch")
        for name in ("output_bytes", "output_capacity", "semantic_change_count"):
            nonnegative_int(row.get(name), f"probe.sample.{name}")
        if row.get("output_bytes") != report.get("output_bytes") or row.get("output_capacity") != report.get("output_capacity"):
            raise ProtocolError("probe sample output size identity mismatch")
        if row.get("semantic_change_count") != cell.get("edits"):
            raise ProtocolError("probe sample semantic change identity mismatch")
        check_oracle(row.get("oracle"), "probe.sample.oracle")
        if allocator:
            if row.get("elapsed_ns") is not None:
                raise ProtocolError("allocator sample elapsed_ns must be null")
            allocation = row.get("allocation")
            if not isinstance(allocation, dict):
                raise ProtocolError("allocator sample has no allocation object")
            for field in ALLOC_FIELDS:
                nonnegative_int(allocation.get(field), f"allocator sample field {field}")
        elif row.get("elapsed_ns") is None or type(row.get("elapsed_ns")) is not int or row.get("elapsed_ns") <= 0:
            raise ProtocolError("native sample elapsed_ns missing or invalid")
        elif row.get("allocation") is not None:
            raise ProtocolError("native sample allocation must be null")


def packet_binary_matches(packet: dict[str, Any], plan: dict[str, Any], role: str) -> None:
    expected = (plan.get("binaries") or {}).get(role) or {}
    observed = packet.get("binary") or {}
    if not isinstance(observed, dict):
        raise ProtocolError(f"{packet.get('packet_id')}: binary is not an object")
    if expected.get("path") and observed.get("path") != expected.get("path"):
        raise ProtocolError(f"{packet.get('packet_id')}: binary path mismatch")
    if expected.get("sha256") and observed.get("sha256") != expected.get("sha256"):
        raise ProtocolError(f"{packet.get('packet_id')}: binary SHA mismatch")
    if expected.get("sha256") and (packet.get("context") or {}).get("binary_sha256") != expected.get("sha256"):
        raise ProtocolError(f"{packet.get('packet_id')}: context binary SHA mismatch")


def check_packet(packet: dict[str, Any], plan: dict[str, Any]) -> None:
    if not isinstance(packet, dict):
        raise ProtocolError("packet is not an object")
    if packet.get("schema") != RUNNER_SCHEMA:
        raise ProtocolError(f"{packet.get('packet_id')}: runner schema mismatch")
    terminal = packet.get("terminal") or {}
    if not isinstance(terminal, dict):
        raise ProtocolError(f"{packet.get('packet_id')}: terminal is not an object")
    if terminal.get("returncode") != 0 or terminal.get("timed_out") or terminal.get("launch_error"):
        raise ProtocolError(f"{packet.get('packet_id')}: child terminal failure")
    if packet.get("validation_error") is not None:
        raise ProtocolError(f"{packet.get('packet_id')}: {packet['validation_error']}")
    context = packet.get("context") or {}
    if not isinstance(context, dict):
        raise ProtocolError(f"{packet.get('packet_id')}: context is not an object")
    role = packet.get("role")
    if not isinstance(role, str):
        raise ProtocolError(f"{packet.get('packet_id')}: role is not a string")
    if role == "aa_baseline":
        binary_role = "native_baseline"
    else:
        binary_role = role
    if binary_role not in {"native_baseline", "native_candidate", "allocator_baseline", "allocator_candidate"}:
        raise ProtocolError(f"{packet.get('packet_id')}: unknown role {role}")
    packet_binary_matches(packet, plan, binary_role)
    rss = packet.get("rss") or {}
    if not isinstance(rss, dict) or type(rss.get("kb")) is not int or rss.get("kb") <= 0:
        raise ProtocolError(f"{packet.get('packet_id')}: RSS packet.rss.kb is missing")
    cell = context.get("cell") or {}
    if not isinstance(cell, dict):
        raise ProtocolError(f"{packet.get('packet_id')}: cell is not an object")
    if cell_id(cell) not in expected_cells():
        raise ProtocolError(f"{packet.get('packet_id')}: invalid cell identity")
    fixture = (plan.get("fixtures") or {}).get(cell.get("case"))
    if not isinstance(fixture, dict):
        raise ProtocolError(f"{packet.get('packet_id')}: fixture binding is missing")
    if context.get("fixture_sha256") != fixture.get("sha256"):
        raise ProtocolError(f"{packet.get('packet_id')}: context fixture SHA mismatch")
    source_role = "baseline" if role == "aa_baseline" or role.endswith("_baseline") else "candidate"
    expected_source_sha = ((plan.get("source") or {}).get(source_role) or {}).get("manifest_sha256")
    if expected_source_sha and context.get("source_manifest_sha256") != expected_source_sha:
        raise ProtocolError(f"{packet.get('packet_id')}: context source manifest SHA mismatch")
    allocator = role.startswith("allocator")
    if role == "aa_baseline":
        allocator = False
    report = packet.get("report")
    if not isinstance(report, dict):
        raise ProtocolError(f"{packet.get('packet_id')}: report is missing")
    check_probe_report(report, cell=cell, fixture=fixture, allocator=allocator)


def child_metric(packet: dict[str, Any], metric: str) -> float:
    report = packet.get("report") or {}
    if metric in {"p50", "p95", "p99"}:
        values = report.get("elapsed_ns")
        if not isinstance(values, list) or not values:
            raise ProtocolError(f"{packet.get('packet_id')}: elapsed_ns unavailable")
        numeric = [float(value) for value in values]
        return max(numeric) if metric == "p99" else percentile(numeric, 0.50 if metric == "p50" else 0.95)
    if metric == "rss_kb":
        value = (packet.get("rss") or {}).get("kb")
        if not isinstance(value, int) or value <= 0:
            raise ProtocolError(f"{packet.get('packet_id')}: packet.rss.kb unavailable")
        return float(value)
    allocation = ((report.get("samples") or [{}])[0]).get("allocation")
    if not isinstance(allocation, dict) or not isinstance(allocation.get(metric), int) or allocation[metric] <= 0:
        raise ProtocolError(f"{packet.get('packet_id')}: allocation.{metric} unavailable")
    return float(allocation[metric])


def pair_metrics(group: dict[str, dict[int, dict[str, Any]]], metric: str, pairs: int, *, seed: int = BOOTSTRAP_SEED, resamples: int = BOOTSTRAP_RESAMPLES) -> dict[str, Any]:
    baseline = group.get("baseline", {})
    candidate = group.get("candidate", {})
    expected = set(range(1, pairs + 1))
    if set(baseline) != expected or set(candidate) != expected:
        raise ProtocolError(f"{metric}: paired children incomplete")
    bvalues = [child_metric(baseline[index], metric) for index in range(1, pairs + 1)]
    cvalues = [child_metric(candidate[index], metric) for index in range(1, pairs + 1)]
    ratios = [ratio(c, b, metric) for c, b in zip(cvalues, bvalues)]
    return {
        "baseline": bvalues,
        "candidate": cvalues,
        "ratios": ratios,
        "bootstrap": bootstrap_median(ratios, seed, resamples),
        "max_ratio": max(ratios),
    }


def grouped_real(packets: list[dict[str, Any]], lane: str) -> dict[str, dict[str, dict[int, dict[str, Any]]]]:
    result: dict[str, dict[str, dict[int, dict[str, Any]]]] = {}
    for packet in packets:
        context = packet.get("context") or {}
        cell = context.get("cell") or {}
        identity = cell_id(cell)
        if identity not in expected_cells():
            raise ProtocolError(f"{packet.get('packet_id')}: invalid real identity")
        variant = context.get("variant")
        pair = context.get("pair")
        if variant not in {"baseline", "candidate"} or not isinstance(pair, int) or pair < 1:
            raise ProtocolError(f"{packet.get('packet_id')}: invalid real pair identity")
        expected_order = "baseline_then_candidate" if pair <= (NATIVE_PAIRS if lane == "native" else ALLOCATOR_PAIRS) // 2 else "candidate_then_baseline"
        if context.get("relative_order") != expected_order:
            raise ProtocolError(f"{packet.get('packet_id')}: relative order mismatch")
        rows = result.setdefault(identity, {}).setdefault(variant, {})
        if pair in rows:
            raise ProtocolError(f"duplicate {lane} packet {identity}/{variant}/pair-{pair}")
        rows[pair] = packet
    return result


def reduce_aa_packets(packets: list[dict[str, Any]], plan: dict[str, Any]) -> dict[str, Any]:
    aa = plan.get("aa_noise") or {}
    cells = aa.get("cells") or expected_cells()
    expected_pairs = set(range(1, NATIVE_PAIRS + 1))
    grouped: dict[str, dict[str, dict[int, dict[str, dict[str, Any]]]]] = {}
    errors: list[str] = []
    for packet in packets:
        try:
            check_packet(packet, plan)
        except ProtocolError as exc:
            errors.append(str(exc))
            continue
        if packet.get("role") != "aa_baseline":
            errors.append(f"{packet.get('packet_id')}: A/A role mismatch")
            continue
        context = packet.get("context") or {}
        cell = (context.get("cell") or {}).get("cell_id")
        order = context.get("aa_order")
        position = context.get("aa_position")
        pair = context.get("pair")
        if cell not in cells or order not in {"AB", "BA"} or position not in {"first", "second"} or not isinstance(pair, int) or pair not in expected_pairs:
            errors.append(f"{packet.get('packet_id')}: invalid A/A context")
            continue
        expected_order = "AB" if pair <= NATIVE_PAIRS // 2 else "BA"
        if order != expected_order:
            errors.append(f"{packet.get('packet_id')}: A/A order mismatch")
            continue
        rows = grouped.setdefault(cell, {}).setdefault(order, {}).setdefault(pair, {})
        if position in rows:
            errors.append(f"{packet.get('packet_id')}: duplicate A/A position")
        rows[position] = packet
    result: dict[str, Any] = {
        "schema": "litchi.performance.0969.aa_readiness.v4",
        "cells": {},
        "errors": errors,
        "bootstrap_seed": BOOTSTRAP_SEED,
        "bootstrap_resamples": BOOTSTRAP_RESAMPLES,
        "pairs_per_cell": NATIVE_PAIRS,
        "upper_ci_limit": CI_LIMIT,
    }
    if errors or set(grouped) != set(cells):
        if set(grouped) != set(cells):
            result["errors"].append("A/A cell coverage is incomplete")
        result["status"] = "incomplete"
        return result
    readiness_fail = False
    for cell in cells:
        cell_result: dict[str, Any] = {}
        for order in ("AB", "BA"):
            rows = grouped[cell].get(order, {})
            expected_order_pairs = {pair for pair in expected_pairs if ("AB" if pair <= NATIVE_PAIRS // 2 else "BA") == order}
            if set(rows) != expected_order_pairs or any(set(row) != {"first", "second"} for row in rows.values()):
                result["errors"].append(f"A/A {cell}/{order}: pair coverage incomplete")
                continue
            order_result: dict[str, Any] = {}
            for metric, quantile in (("p50", 0.50), ("p95", 0.95)):
                ratios: list[float] = []
                for pair in sorted(rows):
                    first = (rows[pair]["first"].get("report") or {}).get("elapsed_ns")
                    second = (rows[pair]["second"].get("report") or {}).get("elapsed_ns")
                    if not isinstance(first, list) or not isinstance(second, list) or len(first) != NATIVE_SAMPLES or len(second) != NATIVE_SAMPLES:
                        result["errors"].append(f"A/A {cell}/{order}/pair-{pair}: sample count mismatch")
                        break
                    ratios.append(ratio(percentile([float(value) for value in second], quantile), percentile([float(value) for value in first], quantile), f"A/A {cell}/{order}/{metric}"))
                if len(ratios) != len(expected_order_pairs):
                    continue
                interval = bootstrap_median(ratios, BOOTSTRAP_SEED, BOOTSTRAP_RESAMPLES)
                symmetric_upper = max(interval["upper"], 1.0 / interval["lower"] if interval["lower"] > 0 else math.inf)
                passed = symmetric_upper <= CI_LIMIT
                readiness_fail = readiness_fail or not passed
                order_result[metric] = {**interval, "ratios": ratios, "symmetric_upper": symmetric_upper, "pass": passed}
            cell_result[order] = order_result
        result["cells"][cell] = cell_result
    if result["errors"]:
        result["status"] = "incomplete"
    else:
        result["status"] = "readiness_fail" if readiness_fail else "pass"
    return result


def reduce_real(packets: list[dict[str, Any]], plan: dict[str, Any]) -> tuple[dict[str, Any], dict[str, Any]]:
    native_packets = [packet for packet in packets if packet.get("role") in {"native_baseline", "native_candidate"}]
    allocator_packets = [packet for packet in packets if packet.get("role") in {"allocator_baseline", "allocator_candidate"}]
    native_grouped = grouped_real(native_packets, "native")
    allocator_grouped = grouped_real(allocator_packets, "allocator")
    native_result: dict[str, Any] = {}
    allocator_result: dict[str, Any] = {}
    for identity in expected_cells():
        group = native_grouped.get(identity)
        if group is None:
            raise ProtocolError(f"native cell missing: {identity}")
        metrics = {metric: pair_metrics(group, metric, NATIVE_PAIRS) for metric in ("p50", "p95", "p99")}
        metrics["noninferior"] = all(metrics[metric]["bootstrap"]["upper"] <= CI_LIMIT for metric in ("p50", "p95"))
        metrics["median_over_1_05_review"] = metrics["p50"]["bootstrap"]["median"] > CI_LIMIT
        metrics["p99_descriptive_flag"] = metrics["p99"]["bootstrap"]["median"] > CI_LIMIT
        native_result[identity] = metrics
        group = allocator_grouped.get(identity)
        if group is None:
            raise ProtocolError(f"allocator cell missing: {identity}")
        metrics = {metric: pair_metrics(group, metric, ALLOCATOR_PAIRS) for metric in ("allocated_bytes", "peak_live_bytes", "allocation_calls", "rss_kb")}
        metrics["per_repeat_guard_pass"] = metrics["allocated_bytes"]["max_ratio"] <= 1.01 and metrics["peak_live_bytes"]["max_ratio"] <= 1.05 and metrics["rss_kb"]["max_ratio"] <= 1.05
        metrics["allocation_calls_median_guard_pass"] = metrics["allocation_calls"]["bootstrap"]["median"] <= 1.01
        allocator_result[identity] = metrics
    return native_result, allocator_result


def decision_from_results(aa: dict[str, Any], native: dict[str, Any] | None, allocator: dict[str, Any] | None) -> dict[str, Any]:
    if aa.get("status") == "incomplete":
        return {"schema": DECISION_SCHEMA, "status": "incomplete", "interpretation": "no_decision", "reason": "A/A packets are missing or invalid"}
    if aa.get("status") == "readiness_fail":
        return {"schema": DECISION_SCHEMA, "status": "readiness_fail", "interpretation": "uninterpretable", "reason": "complete A/A packet exceeded the frozen symmetric noise bound"}
    if aa.get("status") != "pass" or native is None or allocator is None:
        return {"schema": DECISION_SCHEMA, "status": "incomplete", "interpretation": "no_decision", "reason": "formal packets are incomplete"}
    native_pass = all(row.get("noninferior") is True for row in native.values()) and len(native) == 24
    allocator_safety = all(row.get("per_repeat_guard_pass") is True for row in allocator.values()) and len(allocator) == 24
    calls_safety = all(row.get("allocation_calls_median_guard_pass") is True for row in allocator.values()) and len(allocator) == 24
    endpoints = {"docnohf/e32/default", "docfloat/e32/default"}
    requested_benefit = all(allocator.get(cell, {}).get("allocated_bytes", {}).get("bootstrap", {}).get("median", math.inf) <= 0.95 for cell in endpoints)
    native_benefit = any(native.get(cell, {}).get("p50", {}).get("bootstrap", {}).get("median", math.inf) <= 0.95 and native.get(cell, {}).get("p50", {}).get("bootstrap", {}).get("upper", math.inf) < 1.0 for cell in endpoints)
    gates = {"native_noninferiority": native_pass, "allocator_per_repeat_safety": allocator_safety, "allocation_calls_median_safety": calls_safety, "requested_byte_benefit": requested_benefit, "native_benefit": native_benefit}
    qualified = native_pass and allocator_safety and calls_safety and (requested_benefit or native_benefit)
    return {"schema": DECISION_SCHEMA, "status": "qualified" if qualified else "rejected", "interpretation": "complete", "gates": gates, "reason": None if qualified else "one or more hard gates failed"}
