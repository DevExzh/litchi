"""Small, strict custody and report helpers for the 0957 allocator lane.

This module is external evidence tooling. It never executes a benchmark binary.
The report contract is supplied by the frozen plan so a schema revision cannot
be silently accepted by an older reducer.
"""
from __future__ import annotations

import hashlib
import json
import os
import stat
from pathlib import Path
from typing import Any, Iterable

MAX_JSON_BYTES = 2 * 1024 * 1024
PHASES = (
    "setup",
    "ingress",
    "edit_commit",
    "publication",
    "commit_drop",
    "worker_drain",
    "core_combined",
    "core_combined_with_drain",
    "postdrop",
)
METRICS = (
    "allocation_calls",
    "deallocation_calls",
    "reallocation_calls",
    "failed_allocation_calls",
    "allocated_bytes",
    "deallocated_bytes",
    "live_bytes_before",
    "live_bytes_after",
    "peak_live_bytes_before",
    "peak_live_bytes_after",
    "region_peak_live_bytes",
)
NONNEGATIVE = set(METRICS)
ADDITIVE_METRICS = (
    "allocation_calls",
    "deallocation_calls",
    "reallocation_calls",
    "failed_allocation_calls",
    "allocated_bytes",
    "deallocated_bytes",
)
CALIBRATION_COUNTERS = (
    "allocation_calls",
    "deallocation_calls",
    "reallocation_calls",
    "failed_allocation_calls",
)
CONTEXT_LIMIT_FIELDS = (
    "memory",
    "input_bytes",
    "output_bytes",
    "objects",
    "depth",
    "work",
    "execution_workers",
    "execution_in_flight_tasks",
    "execution_in_flight_bytes",
    "execution_min_parallel_bytes",
    "execution_min_task_bytes",
    "root_workers",
    "root_io_concurrency",
    "root_cpu_tasks",
)
RESOURCE_FIELDS = ("memory", "workers", "io_concurrency", "objects")
THREAD_DRAIN_FIELDS = (
    "first_thread_count",
    "final_thread_count",
    "polls",
    "drained",
    "deadline_seconds",
    "poll_limit",
    "poll_sleep_milliseconds",
)
EXPECTED_THREAD_DRAIN = {
    "deadline_seconds": 2,
    "poll_limit": 10_000,
    "poll_sleep_milliseconds": 1,
}


class ContractError(ValueError):
    pass


def _pairs_no_duplicates(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for key, value in pairs:
        if key in result:
            raise ContractError(f"duplicate JSON key: {key}")
        result[key] = value
    return result


def load_json(path: Path, *, max_bytes: int = MAX_JSON_BYTES) -> Any:
    st = path.lstat()
    if stat.S_ISLNK(st.st_mode) or not stat.S_ISREG(st.st_mode):
        raise ContractError(f"JSON input is not a regular file: {path}")
    if st.st_size > max_bytes:
        raise ContractError(f"JSON input exceeds {max_bytes} bytes: {path}")
    raw = path.read_bytes()
    if len(raw) != st.st_size:
        raise ContractError(f"JSON input changed during read: {path}")
    try:
        return json.loads(raw.decode("utf-8"), object_pairs_hook=_pairs_no_duplicates)
    except (UnicodeDecodeError, json.JSONDecodeError, ContractError) as exc:
        raise ContractError(f"invalid JSON {path}: {exc}") from exc


def canonical_json(value: Any) -> bytes:
    return (json.dumps(value, sort_keys=True, separators=(",", ":"), ensure_ascii=True) + "\n").encode()


def sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def file_descriptor(path: Path, *, relative_to: Path | None = None) -> dict[str, Any]:
    st = path.lstat()
    if stat.S_ISLNK(st.st_mode) or not stat.S_ISREG(st.st_mode):
        raise ContractError(f"artifact is not a regular file: {path}")
    name = str(path if relative_to is None else path.relative_to(relative_to))
    return {"path": name, "bytes": st.st_size, "sha256": sha256_file(path)}


def _inside(root: Path, candidate: Path) -> bool:
    try:
        candidate.relative_to(root)
    except ValueError:
        return False
    return True


def packet_path(packet: Path, relative: str, *, must_exist: bool = True) -> Path:
    """Resolve a packet-relative regular file without allowing symlink escape."""
    if not relative or os.path.isabs(relative):
        raise ContractError(f"artifact path must be relative: {relative!r}")
    raw = Path(relative)
    if any(part in ("", ".", "..") for part in raw.parts):
        raise ContractError(f"artifact path is not normalized: {relative!r}")
    root = packet.resolve()
    candidate = (root / raw).resolve(strict=False)
    if not _inside(root, candidate):
        raise ContractError(f"artifact escapes packet: {relative!r}")
    current = root
    for part in raw.parts:
        current = current / part
        if current.is_symlink():
            raise ContractError(f"symlink in packet artifact path: {relative!r}")
    if must_exist:
        st = candidate.lstat()
        if stat.S_ISLNK(st.st_mode) or not stat.S_ISREG(st.st_mode):
            raise ContractError(f"packet artifact is not a regular file: {relative!r}")
    return candidate


def exclusive_json(path: Path, value: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    raw = canonical_json(value)
    with path.open("xb") as stream:
        stream.write(raw)
        stream.flush()
        os.fsync(stream.fileno())


def descriptor_matches(path: Path, descriptor: dict[str, Any]) -> None:
    actual = file_descriptor(path)
    for key in ("bytes", "sha256"):
        if actual[key] != descriptor.get(key):
            raise ContractError(f"{path}: {key} {actual[key]!r} != {descriptor.get(key)!r}")


def require_dict(value: Any, label: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        raise ContractError(f"{label} must be an object")
    return value


def require_str(value: Any, label: str) -> str:
    if not isinstance(value, str) or not value:
        raise ContractError(f"{label} must be a nonempty string")
    return value


def require_int(value: Any, label: str, *, minimum: int = 0) -> int:
    if isinstance(value, bool) or not isinstance(value, int) or value < minimum:
        raise ContractError(f"{label} must be an integer >= {minimum}")
    return value


def expected_children(plan: dict[str, Any]) -> list[dict[str, Any]]:
    schedule = require_dict(plan.get("schedule"), "plan.schedule")
    cases = plan.get("cases")
    if not isinstance(cases, list) or not cases:
        raise ContractError("plan.cases must be a nonempty list")
    widths = schedule.get("workers")
    repeats = require_int(schedule.get("repeats"), "schedule.repeats", minimum=1)
    if widths != [1, 2, 4]:
        raise ContractError(f"schedule.workers must be [1, 2, 4], got {widths!r}")
    if repeats != 3:
        raise ContractError(f"schedule.repeats must be 3, got {repeats!r}")
    result: list[dict[str, Any]] = []
    for case_index, case in enumerate(cases):
        case = require_dict(case, f"plan.cases[{case_index}]")
        case_id = require_str(case.get("case_id"), f"case[{case_index}].case_id")
        for width in widths:
            for repeat in range(repeats):
                result.append({
                    "index": len(result),
                    "child_id": f"{case_id}-w{width}-r{repeat}",
                    "case_id": case_id,
                    "workers": width,
                    "repeat": repeat,
                })
    return result


def validate_plan_schedule(plan: dict[str, Any]) -> list[dict[str, Any]]:
    schedule = require_dict(plan.get("schedule"), "plan.schedule")
    if schedule.get("warmup_iterations") != 2:
        raise ContractError("schedule.warmup_iterations must be 2")
    if schedule.get("measured_samples") != 10:
        raise ContractError("schedule.measured_samples must be 10")
    if schedule.get("deadline_seconds") != 900:
        raise ContractError("schedule.deadline_seconds must be 900")
    if schedule.get("cpu_mask") != [12, 13, 14, 15]:
        raise ContractError("schedule.cpu_mask must be [12, 13, 14, 15]")
    if schedule.get("stop_on_first_failure") is not True:
        raise ContractError("schedule.stop_on_first_failure must be true")
    if schedule.get("no_retries") is not True:
        raise ContractError("schedule.no_retries must be true")
    if schedule.get("process_group") is not True:
        raise ContractError("schedule.process_group must be true")
    case_ids = [case.get("case_id") for case in plan.get("cases", []) if isinstance(case, dict)]
    if schedule.get("case_order") != case_ids:
        raise ContractError("schedule.case_order must match plan case order")
    children = expected_children(plan)
    declared = plan.get("children")
    if declared != children:
        raise ContractError("plan.children is not the canonical 27-child schedule")
    if len(children) != 27:
        raise ContractError(f"expected 27 children, got {len(children)}")
    contract = require_dict(plan.get("report_contract"), "plan.report_contract")
    if contract.get("required_case_fields") != ["admission_cpu_tasks_used"]:
        raise ContractError("plan report contract must require admission_cpu_tasks_used")
    if contract.get("required_row_fields") != [
        "max_setup_core_peak_above_setup_live_before",
        "max_setup_drained_core_peak_above_setup_live_before",
    ]:
        raise ContractError("plan report contract must require both derived peak fields")
    if contract.get("context_limit_fields") != list(CONTEXT_LIMIT_FIELDS):
        raise ContractError("plan report contract context-limit field list is not frozen")
    if contract.get("resource_fields") != list(RESOURCE_FIELDS):
        raise ContractError("plan report contract resource field list is not frozen")
    if contract.get("phase_additive_fields") != list(ADDITIVE_METRICS):
        raise ContractError("plan report contract additive field list is not frozen")
    if contract.get("calibration_serialized_fields") != list(CALIBRATION_COUNTERS):
        raise ContractError("plan report contract calibration field list is not frozen")
    if contract.get("thread_drain") != EXPECTED_THREAD_DRAIN:
        raise ContractError("plan report contract thread-drain configuration is not frozen")
    if plan.get("status") == "authorized":
        for index, case in enumerate(plan.get("cases", [])):
            _validate_authorized_case_oracles(case, f"plan.cases[{index}]")
    return children


def _expected_case(plan: dict[str, Any], case_id: str) -> dict[str, Any]:
    for case in plan.get("cases", []):
        if isinstance(case, dict) and case.get("case_id") == case_id:
            return case
    raise ContractError(f"unknown case id {case_id!r}")


def _validate_authorized_case_oracles(case: Any, label: str) -> None:
    case = require_dict(case, label)
    context_map = case.get("context_limits_by_workers")
    if not isinstance(context_map, dict) or set(context_map) != {"1", "2", "4"}:
        raise ContractError(f"{label}.context_limits_by_workers must contain workers 1,2,4")
    for worker, values in context_map.items():
        _check_context_limits(values, f"{label}.context_limits_by_workers[{worker!r}]")
    admission_map = case.get("admission_by_workers")
    if not isinstance(admission_map, dict) or set(admission_map) != {"1", "2", "4"}:
        raise ContractError(f"{label}.admission_by_workers must contain workers 1,2,4")
    for worker, values in admission_map.items():
        values = require_dict(values, f"{label}.admission_by_workers[{worker!r}]")
        for field in (
            "admission_expected_worker_permits",
            "admission_observed_worker_permits",
            "admission_cpu_tasks_used",
        ):
            require_int(values.get(field), f"{label}.admission_by_workers[{worker!r}].{field}")
    for field in ("source_bytes", "output_bytes", "update_count", "sink_capacity_bytes", "sink_max_write_bytes"):
        require_int(case.get(field), f"{label}.{field}")
    for field in ("source_sha256", "output_sha256", "semantic_sha256"):
        _require_sha(case.get(field), f"{label}.{field}")


def _require_sha(value: Any, label: str) -> str:
    if not isinstance(value, str) or len(value) != 64:
        raise ContractError(f"{label} must be a 64-character SHA-256 string")
    try:
        int(value, 16)
    except ValueError as exc:
        raise ContractError(f"{label} is not hexadecimal SHA-256") from exc
    return value


def _check_context_limits(value: Any, label: str) -> dict[str, Any]:
    value = require_dict(value, label)
    if set(value) != set(CONTEXT_LIMIT_FIELDS):
        raise ContractError(f"{label} must contain exactly the context-limit fields")
    for field in CONTEXT_LIMIT_FIELDS:
        require_int(value.get(field), f"{label}.{field}")
    return value


def _check_resource_snapshot(value: Any, label: str) -> dict[str, Any]:
    value = require_dict(value, label)
    if set(value) != set(RESOURCE_FIELDS):
        raise ContractError(f"{label} must contain exactly the resource fields")
    for field in RESOURCE_FIELDS:
        if require_int(value.get(field), f"{label}.{field}") != 0:
            raise ContractError(f"{label}.{field} retained a managed resource")
    return value


def _check_thread_drain(value: Any, label: str, *, require_first_one: bool) -> dict[str, Any]:
    value = require_dict(value, label)
    if set(value) != set(THREAD_DRAIN_FIELDS):
        raise ContractError(f"{label} must contain exactly the thread-drain fields")
    for field in ("first_thread_count", "final_thread_count", "polls", "deadline_seconds", "poll_limit", "poll_sleep_milliseconds"):
        require_int(value.get(field), f"{label}.{field}")
    if value.get("drained") is not True:
        raise ContractError(f"{label}.drained must be true")
    if require_first_one and value.get("first_thread_count") != 1:
        raise ContractError(f"{label}.first_thread_count must be 1")
    if value.get("final_thread_count") != 1 or value.get("polls") < 1:
        raise ContractError(f"{label} did not finish at one thread")
    if value.get("polls") > value.get("poll_limit"):
        raise ContractError(f"{label}.polls exceeds its configured poll limit")
    for field, expected in EXPECTED_THREAD_DRAIN.items():
        if value.get(field) != expected:
            raise ContractError(f"{label}.{field} does not match the frozen observer configuration")
    return value


def _check_sample(sample: Any, label: str) -> None:
    sample = require_dict(sample, label)
    if sample.get("status") != "measured":
        raise ContractError(f"{label}.status is not measured")
    if sample.get("scope") != "operation_global_system_allocator":
        raise ContractError(f"{label}.scope is not operation_global_system_allocator")
    for key in METRICS:
        value = sample.get(key)
        if value is None or isinstance(value, bool) or not isinstance(value, int) or value < 0:
            raise ContractError(f"{label}.{key} must be a nonnegative integer")
    if sample["region_peak_live_bytes"] < max(sample["live_bytes_before"], sample["live_bytes_after"]):
        raise ContractError(f"{label}.region_peak_live_bytes is below live boundary")
    if sample["peak_live_bytes_before"] < sample["live_bytes_before"]:
        raise ContractError(f"{label}.peak_live_bytes_before is below live boundary")
    if sample["peak_live_bytes_after"] < max(
        sample["peak_live_bytes_before"],
        sample["live_bytes_after"],
        sample["region_peak_live_bytes"],
    ):
        raise ContractError(f"{label}.peak_live_bytes_after does not cover its observed maxima")
    if sample["failed_allocation_calls"] != 0:
        raise ContractError(f"{label}.failed_allocation_calls must be zero")


def _check_adjacent(left: dict[str, Any], right: dict[str, Any], label: str) -> None:
    if left["live_bytes_after"] != right["live_bytes_before"]:
        raise ContractError(f"{label} live-byte boundaries are not adjacent")


def _check_combined(
    row: dict[str, Any],
    name: str,
    parts: tuple[str, ...],
) -> None:
    combined = row[name]
    _check_sample(combined, f"row.{name}")
    first = row[parts[0]]
    last = row[parts[-1]]
    if combined["live_bytes_before"] != first["live_bytes_before"]:
        raise ContractError(f"{name}.live_bytes_before is not the first phase boundary")
    if combined["live_bytes_after"] != last["live_bytes_after"]:
        raise ContractError(f"{name}.live_bytes_after is not the last phase boundary")
    if combined["peak_live_bytes_before"] != first["peak_live_bytes_before"]:
        raise ContractError(f"{name}.peak_live_bytes_before is not the first phase maximum")
    expected_peak_after = max(row[part]["peak_live_bytes_after"] for part in parts)
    if combined["peak_live_bytes_after"] != expected_peak_after:
        raise ContractError(f"{name}.peak_live_bytes_after is not the exact phase maximum")
    if combined["region_peak_live_bytes"] != max(row[part]["region_peak_live_bytes"] for part in parts):
        raise ContractError(f"{name}.region_peak_live_bytes is not the exact phase maximum")
    for metric in ADDITIVE_METRICS:
        expected = sum(row[part][metric] for part in parts)
        if combined[metric] != expected:
            raise ContractError(f"{name}.{metric} conservation failed")


def _check_context_and_admission(
    case: dict[str, Any],
    row: dict[str, Any],
    expected_case: dict[str, Any],
    workers: int,
    *,
    label: str,
) -> None:
    case_context = _check_context_limits(case.get("context_limits"), f"{label}.case.context_limits")
    row_context = _check_context_limits(row.get("context_limits"), f"{label}.row.context_limits")
    if row_context != case_context:
        raise ContractError(f"{label} row context limits differ from case context limits")
    admission_map = expected_case.get("admission_by_workers")
    context_map = expected_case.get("context_limits_by_workers")
    if isinstance(admission_map, dict) and isinstance(context_map, dict):
        admission = require_dict(admission_map.get(str(workers)), f"{label}.admission_by_workers")
        expected_context = _check_context_limits(
            context_map.get(str(workers)), f"{label}.context_limits_by_workers"
        )
        if case_context != expected_context:
            raise ContractError(f"{label} context limits do not match the worker oracle")
        for field in (
            "admission_expected_worker_permits",
            "admission_observed_worker_permits",
            "admission_cpu_tasks_used",
        ):
            if case.get(field) != admission.get(field):
                raise ContractError(f"{label}.case.{field} does not match the admission oracle")
        if row.get("expected_worker_permits") != admission["admission_expected_worker_permits"]:
            raise ContractError(f"{label}.row.expected_worker_permits does not match the admission oracle")
        if row.get("admitted_worker_permits") != admission["admission_observed_worker_permits"]:
            raise ContractError(f"{label}.row.admitted_worker_permits does not match the admission oracle")
    else:
        for field in (
            "admission_expected_worker_permits",
            "admission_observed_worker_permits",
            "admission_cpu_tasks_used",
        ):
            require_int(case.get(field), f"{label}.case.{field}")
        require_int(row.get("expected_worker_permits"), f"{label}.row.expected_worker_permits")
        require_int(row.get("admitted_worker_permits"), f"{label}.row.admitted_worker_permits")
    if row.get("expected_cpu_tasks") != expected_case.get("expected_cpu_tasks"):
        raise ContractError(f"{label}.row.expected_cpu_tasks does not match the case oracle")
    if row.get("cpu_tasks_used") != expected_case.get("expected_cpu_tasks"):
        raise ContractError(f"{label}.row.cpu_tasks_used does not match the case oracle")
    if case.get("admission_cpu_tasks_used") != expected_case.get("expected_cpu_tasks"):
        raise ContractError(f"{label}.case.admission_cpu_tasks_used does not match expected CPU tasks")


def _check_artifacts_and_sink(case: dict[str, Any], row: dict[str, Any], *, label: str) -> None:
    for field in ("source_sha256", "output_sha256", "semantic_sha256"):
        if row.get(field) != case.get(field):
            raise ContractError(f"{label}.{field} does not match the case artifact identity")
    if row.get("output_bytes") != case.get("output_bytes"):
        raise ContractError(f"{label}.output_bytes does not match the case oracle")
    if row.get("output_bytes_used") != row.get("output_bytes"):
        raise ContractError(f"{label}.output_bytes_used does not equal output_bytes")
    if row.get("output_bytes") > case.get("sink_capacity_bytes"):
        raise ContractError(f"{label}.output_bytes exceeds sink capacity")
    if not isinstance(row.get("sink_write_calls"), int) or row["sink_write_calls"] < 1:
        raise ContractError(f"{label}.sink_write_calls must be positive")
    if not isinstance(row.get("sink_largest_write"), int) or not 1 <= row["sink_largest_write"] <= case.get("sink_max_write_bytes"):
        raise ContractError(f"{label}.sink_largest_write exceeds the frozen sink bound")
    _check_resource_snapshot(row.get("resources_after_core"), f"{label}.resources_after_core")


def validate_report(report: Any, plan: dict[str, Any], child: dict[str, Any]) -> dict[str, Any]:
    report = require_dict(report, "report")
    contract = require_dict(plan.get("report_contract"), "plan.report_contract")
    if report.get("schema") != contract.get("schema"):
        raise ContractError(f"report schema {report.get('schema')!r} != {contract.get('schema')!r}")
    if report.get("status") != "completed":
        raise ContractError("report.status is not completed")
    expected_binary = require_dict(plan.get("binary"), "plan.binary")
    binary = require_dict(report.get("binary_identity"), "report.binary_identity")
    if binary.get("sha256") != expected_binary.get("sha256") or binary.get("bytes") != expected_binary.get("bytes"):
        raise ContractError("report binary identity does not match frozen binary")
    for key in ("allocator_identity", "instrumentation_identity", "counter_revision"):
        expected = contract.get(key)
        if expected is not None and report.get(key) != expected:
            raise ContractError(f"report.{key} does not match frozen contract")
    for key in ("allocation_scope", "thread_observation_scope", "timing_scope", "performance_claim", "peak_semantics"):
        expected = contract.get(key)
        if expected is not None and report.get(key) != expected:
            raise ContractError(f"report.{key} does not match frozen contract")

    required_case_fields = contract.get("required_case_fields", [])
    required_row_fields = contract.get("required_row_fields", [])
    if required_case_fields != ["admission_cpu_tasks_used"]:
        raise ContractError("report contract required_case_fields is not the frozen list")
    if required_row_fields != [
        "max_setup_core_peak_above_setup_live_before",
        "max_setup_drained_core_peak_above_setup_live_before",
    ]:
        raise ContractError("report contract required_row_fields is not the frozen list")

    calibration = require_dict(report.get("calibration"), "report.calibration")
    if calibration.get("status") != "pass" or calibration.get("no_allocator_callbacks") is not True:
        raise ContractError("allocator calibration did not pass")
    if calibration.get("polls") != 100:
        raise ContractError("allocator calibration did not perform exactly 100 polls")
    if calibration.get("thread_count_first") != 1 or calibration.get("thread_count_last") != 1:
        raise ContractError("allocator calibration did not observe one thread")
    if calibration.get("thread_scope") != report.get("thread_observation_scope"):
        raise ContractError("allocator calibration thread scope differs from report scope")
    for field in CALIBRATION_COUNTERS:
        if calibration.get(field) != 0:
            raise ContractError(f"allocator calibration {field} was not zero")

    case = require_dict(report.get("case"), "report.case")
    missing_case_fields = [field for field in required_case_fields if field not in case]
    if missing_case_fields:
        raise ContractError(f"report.case is missing required fields: {missing_case_fields!r}")
    expected_case = _expected_case(plan, child["case_id"])
    field_map = {
        "shape": "shape",
        "workload": "workload",
        "workers": "workers",
        "source_bytes": "source_bytes",
        "source_sha256": "source_sha256",
        "output_bytes": "output_bytes",
        "output_sha256": "output_sha256",
        "semantic_sha256": "semantic_sha256",
        "update_count": "update_count",
        "expected_cpu_tasks": "expected_cpu_tasks",
        "sink_capacity_bytes": "sink_capacity_bytes",
        "sink_max_write_bytes": "sink_max_write_bytes",
        "warmup_iterations": "warmup_iterations",
        "measured_samples": "measured_samples",
    }
    for report_key, case_key in field_map.items():
        expected = expected_case.get(case_key)
        if expected is not None and case.get(report_key) != expected:
            raise ContractError(f"report.case.{report_key} does not match frozen case oracle")
    for field in ("source_bytes", "output_bytes", "update_count", "expected_cpu_tasks", "sink_capacity_bytes", "sink_max_write_bytes"):
        require_int(case.get(field), f"report.case.{field}")
    for field in ("source_sha256", "output_sha256", "semantic_sha256"):
        value = case.get(field)
        if plan.get("status") == "authorized":
            _require_sha(value, f"report.case.{field}")
        elif not isinstance(value, str) or not value:
            raise ContractError(f"report.case.{field} must be a nonempty string")
    if case.get("workers") != child["workers"]:
        raise ContractError("report.case.workers does not match schedule")
    expected_id = f"{child['case_id']}-workers{child['workers']}"
    if case.get("id") != expected_id:
        raise ContractError(f"report.case.id {case.get('id')!r} != {expected_id!r}")
    schedule = require_dict(plan.get("schedule"), "plan.schedule")
    expected_warmup = schedule.get("warmup_iterations")
    expected_samples = schedule.get("measured_samples")
    if case.get("warmup_iterations") != expected_warmup or case.get("measured_samples") != expected_samples:
        raise ContractError("report case sample schedule does not match frozen plan")
    if case.get("preflight_actual_commit_inverse_verified") is not True:
        raise ContractError("actual commit inverse preflight was not verified")
    rows = case.get("rows")
    if not isinstance(rows, list) or len(rows) != expected_samples:
        raise ContractError("report.case.rows count does not match frozen plan")
    phases = contract.get("phases", list(PHASES))
    if phases != list(PHASES):
        raise ContractError(f"unsupported phase contract: {phases!r}")

    for index, row in enumerate(rows):
        row = require_dict(row, f"report.case.rows[{index}]")
        missing_row_fields = [field for field in required_row_fields if field not in row]
        if missing_row_fields:
            raise ContractError(f"report.case.rows[{index}] is missing required fields: {missing_row_fields!r}")
        if row.get("sample") != index:
            raise ContractError(f"row sample index {row.get('sample')!r} != {index}")
        if row.get("fixture_inverse_restored_exact_source") is not True:
            raise ContractError(f"row {index} inverse oracle did not pass")
        if row.get("workers_requested") != child["workers"]:
            raise ContractError(f"row {index} workers_requested mismatch")
        _check_context_and_admission(
            case,
            row,
            expected_case,
            child["workers"],
            label=f"report.case.rows[{index}]",
        )
        for field in (
            "expected_cpu_tasks",
            "cpu_tasks_used",
            "output_bytes",
            "output_bytes_used",
            "sink_write_calls",
            "sink_largest_write",
        ):
            require_int(row.get(field), f"report.case.rows[{index}].{field}")
        _check_artifacts_and_sink(case, row, label=f"report.case.rows[{index}]")
        for peak_field in required_row_fields:
            require_int(row.get(peak_field), f"report.case.rows[{index}].{peak_field}")
        if row.get("source_sha256") != case.get("source_sha256") or row.get("output_sha256") != case.get("output_sha256"):
            raise ContractError(f"row {index} artifact identity mismatch")
        if row.get("semantic_sha256") != case.get("semantic_sha256"):
            raise ContractError(f"row {index} semantic artifact identity mismatch")
        for phase in phases:
            _check_sample(row.get(phase), f"report.case.rows[{index}].{phase}")
        _check_thread_drain(
            row.get("pre_sample_threads"),
            f"report.case.rows[{index}].pre_sample_threads",
            require_first_one=True,
        )
        _check_thread_drain(
            row.get("post_core_threads"),
            f"report.case.rows[{index}].post_core_threads",
            require_first_one=False,
        )
        for left, right in zip(
            ("setup", "ingress", "edit_commit", "publication", "commit_drop"),
            ("ingress", "edit_commit", "publication", "commit_drop", "worker_drain"),
        ):
            _check_adjacent(row[left], row[right], f"row {index} {left}->{right}")
        _check_combined(
            row,
            "core_combined",
            ("ingress", "edit_commit", "publication", "commit_drop"),
        )
        _check_combined(
            row,
            "core_combined_with_drain",
            ("ingress", "edit_commit", "publication", "commit_drop", "worker_drain"),
        )
        baseline = row["setup"]["live_bytes_before"]
        core_peak = max(
            row["setup"]["region_peak_live_bytes"],
            row["core_combined"]["region_peak_live_bytes"],
        )
        drained_peak = max(
            row["setup"]["region_peak_live_bytes"],
            row["core_combined_with_drain"]["region_peak_live_bytes"],
        )
        if core_peak < baseline or drained_peak < baseline:
            raise ContractError(f"row {index} derived peak baseline underflowed")
        if row["max_setup_core_peak_above_setup_live_before"] != core_peak - baseline:
            raise ContractError(f"row {index} max_setup_core_peak_above_setup_live_before is not exact")
        if row["max_setup_drained_core_peak_above_setup_live_before"] != drained_peak - baseline:
            raise ContractError(f"row {index} max_setup_drained_core_peak_above_setup_live_before is not exact")
    return report

def finite_metric(sample: dict[str, Any], metric: str) -> int:
    value = sample.get(metric)
    if isinstance(value, bool) or not isinstance(value, int) or value < 0:
        raise ContractError(f"metric {metric} is not finite")
    return value


def mean(values: Iterable[int]) -> float:
    values = list(values)
    if not values:
        raise ContractError("cannot summarize empty values")
    return sum(values) / len(values)
