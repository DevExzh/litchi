#!/usr/bin/env python3
"""Fail-closed offline reducer for the 0978 attribution packet contract.

The reducer deliberately owns no capture, decoder, raw scanner, or lifecycle
implementation.  It consumes immutable sidecars produced by those helpers,
reparses their sample ledgers, and refuses to produce a cohort result when a
packet, sidecar, or claim disagrees.  ``D``/``M``/``R`` describe sampled
instruction periods; ``R`` is explicitly ``marker ancestry not observed``.
Target-source owner bounds are reported separately from all-DSO physical leaf
and inclusive marker views.
"""
from __future__ import annotations

import argparse
import hashlib
import importlib.util
import json
import math
import re
import statistics
import sys
from collections import Counter
from pathlib import Path
from typing import Any, Iterable, Mapping, Sequence


SCHEMA = "litchi.performance.0978.instruction_profile_analysis.v1"
PACKET_SCHEMA = "litchi.performance.0978.child_packet.v1"
MANIFEST_SCHEMA = "litchi.performance.0978.capture_manifest.v1"
CHILDREN = 12
PERIOD = 1_000_000
DEFAULT_MARKER = "doc_retention_probe_0942_cfb_profile::profiled_measured_case"
PLAN_SCHEMA_PREFIX = "litchi.performance.0978.plan"

# The retained 0977 decoder summary checker predates the 0978 disjoint
# lifecycle ledger.  These fields describe physical-leaf/unknown ownership
# and are therefore checked by the 0978 attribution ledger below.  Feeding
# them through the old checker would compare two different ownership
# definitions (and reject valid mapped-DSO/unknown rows).
LEGACY_OWNER_SUMMARY_KEYS = frozenset({
    "physical_leaf_owner_period",
    "physical_leaf_unknown_period",
    "physical_leaf_unknown_samples",
    "marker_unknown_physical_leaf_samples",
    "marker_no_resolved_frame_samples",
    "physical_leaf_classified_period",
    "physical_leaf_conservation",
    "physical_leaf_scope",
    "deepest_resolved_proxy_period",
    "inclusive_marker_ancestor_period",
})

# New 0978 summary fields remain explicit and are validated by
# _compare_packet_claims/ledger conservation.  Unknown additions fail closed
# instead of disappearing through the compatibility projection.
SUMMARY_0978_KEYS = frozenset({
    "schema",
    "marker_known_owner_period",
    "target_known_owner_period",
    "target_source_eligible_period",
    "inclusive_owner_period",
    "lifecycle_unknown_period",
    "shim_log_sha256",
})


class ReducerFailure(ValueError):
    """A fail-closed input or conservation failure."""

    def __init__(self, code: str, message: str, **details: Any):
        super().__init__(f"{code}: {message}")
        self.code = code
        self.message = message
        self.details = details

    def as_dict(self) -> dict[str, Any]:
        return {"status": "failed", "code": self.code, "message": self.message, "details": self.details}


def _fail(code: str, message: str, **details: Any) -> None:
    raise ReducerFailure(code, message, **details)


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with Path(path).open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def file_binding(path: Path) -> dict[str, Any]:
    path = Path(path).absolute()
    if not path.is_file():
        _fail("input_missing", "bound file is missing", path=str(path))
    return {"path": str(path), "bytes": path.stat().st_size, "sha256": sha256(path)}


def check_binding(row: Mapping[str, Any], label: str) -> Path:
    if not isinstance(row, Mapping):
        _fail("binding_shape", f"{label} is not an object")
    path_text = row.get("path")
    if not isinstance(path_text, str) or not path_text:
        _fail("binding_shape", f"{label}.path is missing")
    path = Path(path_text).absolute()
    if not path.is_file():
        _fail("input_missing", f"{label} path is missing", path=str(path))
    actual = file_binding(path)
    for key in ("bytes", "sha256"):
        if row.get(key) != actual[key]:
            _fail("binding_mismatch", f"{label}.{key} differs from file", expected=row.get(key), actual=actual[key])
    return path


def load_json(path: Path, label: str) -> dict[str, Any]:
    try:
        value = json.loads(Path(path).read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as exc:
        _fail("json_input", f"cannot load {label}: {exc}")
    if not isinstance(value, dict):
        _fail("json_shape", f"{label} is not an object")
    return value


def _finite_number(value: Any, label: str) -> float:
    if isinstance(value, bool) or not isinstance(value, (int, float)) or not math.isfinite(float(value)):
        _fail("number", f"{label} is not finite", value=value)
    return float(value)


def _positive_int(value: Any, label: str) -> int:
    if isinstance(value, bool) or not isinstance(value, int) or value <= 0:
        _fail("integer", f"{label} is not a positive integer", value=value)
    return value


def _nonnegative_int(value: Any, label: str) -> int:
    if isinstance(value, bool) or not isinstance(value, int) or value < 0:
        _fail("integer", f"{label} is not a nonnegative integer", value=value)
    return value


def _same(left: Any, right: Any) -> bool:
    return left == right


def _first_mapping(obj: Mapping[str, Any], keys: Sequence[str]) -> Any:
    found: list[tuple[str, Any]] = [(key, obj[key]) for key in keys if key in obj]
    if not found:
        return None
    first = found[0][1]
    for key, value in found[1:]:
        if value != first:
            _fail("alias_conflict", "sidecar aliases disagree", aliases=[name for name, _ in found], conflicting=key)
    return first


def _artifact_map(packet: Mapping[str, Any]) -> dict[str, dict[str, Any]]:
    merged: dict[str, dict[str, Any]] = {}
    # The packet contract intentionally has one canonical artifact namespace.
    # ``runner_bindings`` is accepted only for the inherited 0977 custody rows;
    # the seven 0978 inputs below must be present under ``artifacts`` with
    # these exact names.
    for source_name in ("artifacts", "runner_bindings"):
        source = packet.get(source_name)
        if source is None:
            continue
        if not isinstance(source, Mapping):
            _fail("artifact_shape", f"{packet.get('packet_id')}: {source_name} is not an object")
        for name, row in source.items():
            if not isinstance(row, Mapping):
                _fail("artifact_shape", f"{packet.get('packet_id')}: artifact {name} is not a binding")
            checked = {"path": row.get("path"), "bytes": row.get("bytes"), "sha256": row.get("sha256")}
            check_binding(checked, f"{packet.get('packet_id')}.{source_name}.{name}")
            previous = merged.get(str(name))
            if previous is not None and previous != checked:
                _fail("artifact_conflict", f"{packet.get('packet_id')}: artifact aliases conflict", name=name)
            merged[str(name)] = checked
    if not merged:
        _fail("artifact_missing", f"{packet.get('packet_id')}: no artifact bindings")
    required = (
        "corrected_script", "corrected_log", "perf_mmap_dump", "perf_evlist",
        "raw_stack", "lifecycle_ledger", "attribution_samples",
    )
    missing = [name for name in required if name not in merged]
    if missing:
        _fail("artifact_missing", f"{packet.get('packet_id')}: missing canonical artifacts", missing=missing)
    return merged


def _read_artifact(artifacts: Mapping[str, Mapping[str, Any]], name: str, label: str) -> tuple[Path, dict[str, Any] | None]:
    row = artifacts.get(name)
    if row is None:
        return Path(), None
    path = check_binding(row, label)
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except (UnicodeDecodeError, json.JSONDecodeError):
        return path, None
    if not isinstance(value, dict):
        _fail("sidecar_shape", f"{label} is not a JSON object")
    return path, value


def _validate_plan(plan_path: Path) -> dict[str, Any]:
    plan_path = Path(plan_path).absolute()
    plan = load_json(plan_path, "plan")
    if not str(plan.get("schema", "")).startswith(PLAN_SCHEMA_PREFIX):
        _fail("plan_contract", "plan schema is not the 0978 plan schema", schema=plan.get("schema"))
    # The plan-binding module owns the complete pre-capture custody/oracle
    # contract.  Reusing that validator keeps this reducer from accepting a
    # packet whose seven analysis sidecars happen to hash correctly while the
    # native command, report, current-perf receipts, or helper set drifted.
    try:
        binding_module = _load_sibling_module("litchi_plan_binding_0978_reducer", "plan_binding_0978_v1.py")
        result = binding_module.validate_plan(plan_path)
    except Exception as exc:
        _fail("plan_contract", f"full 0978 plan binding validation failed: {exc}")
    if not isinstance(result, Mapping) or result.get("status") != "pass":
        _fail("plan_contract", "full 0978 plan binding validator did not return pass")
    return plan


def _rows_from_sidecar(obj: Mapping[str, Any], label: str) -> list[Mapping[str, Any]]:
    value = _first_mapping(obj, ("sample_ledger", "samples", "rows"))
    if value is None and isinstance(obj.get("ledger"), Mapping):
        value = _first_mapping(obj["ledger"], ("sample_ledger", "samples", "rows"))
    if not isinstance(value, list):
        _fail("sample_ledger_missing", f"{label} does not retain a sample ledger")
    if any(not isinstance(row, Mapping) for row in value):
        _fail("sample_ledger_shape", f"{label} contains a non-object sample")
    return list(value)


def _rows_from_attribution(obj: Mapping[str, Any], name: str) -> list[Mapping[str, Any]]:
    """Read one canonical full per-sample stream from attribution_samples."""
    value = obj.get(name)
    if value is None and isinstance(obj.get("streams"), Mapping):
        value = obj["streams"].get(name)
    if not isinstance(value, list):
        _fail("sample_ledger_missing", f"attribution_samples.{name} is missing")
    if any(not isinstance(row, Mapping) for row in value):
        _fail("sample_ledger_shape", f"attribution_samples.{name} contains a non-object sample")
    return list(value)


def _frame_dso(frame: Mapping[str, Any]) -> str | None:
    for key in ("dso", "path", "file", "object"):
        value = frame.get(key)
        if isinstance(value, str) and value:
            return value
    return None


def _frame_symbol(frame: Mapping[str, Any]) -> str | None:
    for key in ("symbol_base", "symbol", "name", "function"):
        value = frame.get(key)
        if isinstance(value, str) and value:
            return value
    return None


def _frame_resolved(frame: Mapping[str, Any]) -> bool:
    if frame.get("resolved") is False:
        return False
    return _frame_dso(frame) is not None and _frame_symbol(frame) is not None


def _owner_key(frame: Mapping[str, Any]) -> str | None:
    explicit = frame.get("owner") or frame.get("physical_owner")
    if isinstance(explicit, str) and explicit:
        return explicit
    dso, symbol = _frame_dso(frame), _frame_symbol(frame)
    if dso and symbol:
        return f"{dso}::{symbol}"
    return None


def _marker_match(frames: Sequence[Mapping[str, Any]], marker: str, marker_dso: str) -> bool:
    for frame in frames:
        if _frame_dso(frame) == marker_dso and _frame_symbol(frame) == marker:
            return True
    return False


def _normalise_sample(row: Mapping[str, Any], index: int, label: str, *, marker: str, marker_dso: str, allow_summary_claims: bool = False, allow_unframed_claims: bool = False) -> dict[str, Any]:
    raw_index = row.get("index", row.get("sample_index"))
    if raw_index != index:
        _fail("sample_identity", f"{label}: sample index is not contiguous", expected=index, actual=raw_index)
    period = _positive_int(row.get("period"), f"{label}[{index}].period")
    time_value = row.get("time_ns", row.get("time"))
    if isinstance(time_value, bool) or not isinstance(time_value, int) or time_value < 0:
        _fail("sample_identity", f"{label}[{index}].time_ns is invalid", value=time_value)
    pid = row.get("pid")
    tid = row.get("tid", pid)
    cpu = row.get("cpu")
    for name, value in (("pid", pid), ("tid", tid), ("cpu", cpu)):
        if value is None or isinstance(value, bool) or not isinstance(value, int) or value < 0:
            _fail("sample_identity", f"{label}[{index}].{name} is invalid", value=value)
    frames_supplied = "frames" in row or "callchain" in row
    frames_value = row.get("frames", row.get("callchain", []))
    if frames_value is None:
        frames_value = []
    if not isinstance(frames_value, list) or any(not isinstance(frame, Mapping) for frame in frames_value):
        _fail("frame_shape", f"{label}[{index}] frames are not objects")
    frames = [dict(frame) for frame in frames_value]
    marker_actual = _marker_match(frames, marker, marker_dso)
    if not frames_supplied and allow_summary_claims and isinstance(row.get("marker"), bool):
        marker_actual = bool(row["marker"])
        if marker_actual:
            marker_info = row.get("marker_frame")
            if isinstance(marker_info, Mapping):
                if marker_info.get("dso_path", marker_info.get("dso")) != marker_dso or marker_info.get("symbol") != marker:
                    _fail("marker_mismatch", f"{label}[{index}] marker summary has wrong DSO/symbol")
            # A lifecycle ledger's marker boolean has already been validated by
            # its owner; the full attribution sidecar is checked below.
    elif "marker" in row and row.get("marker") != marker_actual:
        _fail("marker_mismatch", f"{label}[{index}] marker claim differs from frame DSO/symbol", claimed=row.get("marker"), actual=marker_actual)
    marker_dso_claim = row.get("marker_dso")
    if marker_actual and marker_dso_claim is not None and marker_dso_claim != marker_dso:
        _fail("marker_mismatch", f"{label}[{index}] marker DSO claim differs", claimed=marker_dso_claim, expected=marker_dso)
    resolved = [frame for frame in frames if _frame_resolved(frame)]
    physical_owner = _owner_key(frames[0]) if frames and _frame_resolved(frames[0]) else None
    inclusive = sorted({_owner_key(frame) for frame in resolved if _owner_key(frame) is not None})
    if not frames_supplied and allow_summary_claims:
        if not allow_unframed_claims and (
            row.get("physical_owner") is not None
            or row.get("inclusive_owners") is not None
            or row.get("inclusive_owner_period") is not None
            or row.get("marker") is True
        ):
            _fail("owner_proof", f"{label}[{index}] claims attribution without retained frames")
        explicit_owner = row.get("physical_owner")
        if explicit_owner is not None and (not isinstance(explicit_owner, str) or not explicit_owner):
            _fail("owner_shape", f"{label}[{index}] physical_owner is invalid")
        physical_owner = explicit_owner
        explicit_inclusive = row.get("inclusive_owners", row.get("inclusive_owner_period"))
        if explicit_inclusive is not None:
            if isinstance(explicit_inclusive, Mapping):
                inclusive = sorted(str(key) for key in explicit_inclusive)
            elif isinstance(explicit_inclusive, list) and all(isinstance(value, str) for value in explicit_inclusive):
                inclusive = sorted(set(explicit_inclusive))
            else:
                _fail("owner_shape", f"{label}[{index}] inclusive owners are invalid")
    target_eligible = row.get("target_source_eligible")
    if target_eligible is not None and not isinstance(target_eligible, bool):
        _fail("target_eligibility", f"{label}[{index}] target_source_eligible is not boolean")
    return {
        "index": index,
        "period": period,
        "time_ns": time_value,
        "pid": pid,
        "tid": tid,
        "cpu": cpu,
        "frames": frames,
        "marker": marker_actual,
        "physical_owner": physical_owner,
        "inclusive_owners": inclusive,
        "target_source_eligible": target_eligible,
        "phase": row.get("phase"),
        "mode": row.get("mode"),
        "unknown": row.get("unknown"),
        "unknown_reason": row.get("unknown_reason"),
        "raw_identity": row.get("raw_identity"),
    }


def _parse_structured_rows(obj: Mapping[str, Any], label: str, *, marker: str, marker_dso: str, allow_summary_claims: bool = False, allow_unframed_claims: bool = False) -> list[dict[str, Any]]:
    rows = _rows_from_sidecar(obj, label)
    return [_normalise_sample(row, index, label, marker=marker, marker_dso=marker_dso, allow_summary_claims=allow_summary_claims, allow_unframed_claims=allow_unframed_claims) for index, row in enumerate(rows, 1)]


_TEXT_HEADER = re.compile(
    r"^\s*(?:(?P<comm>\S+)\s+)?(?P<pid>\d+)(?:/(?P<tid>\d+))?"
    r"\s+\[(?P<cpu>\d+)\]\s+(?P<time>\d+(?:\.\d+)?):\s+"
    r"(?P<period>\d+)\s+(?P<event>.+?):\s*$"
)
_TEXT_FRAME = re.compile(r"^\s*(?P<ip>[0-9a-fA-F]+)\s+(?P<symbol>.*?)\s+\((?P<dso>.*)\)\s*$")


def _parse_perf_ip_token(token: str, label: str) -> int:
    """Parse perf's frame-address token as hexadecimal, including digits-only tokens."""
    if not isinstance(token, str) or re.fullmatch(r"[0-9a-fA-F]+", token) is None:
        _fail("corrected_frame", f"{label} is not a hexadecimal perf IP", value=token)
    return int(token, 16)


def _parse_corrected_text(path: Path, *, marker: str, marker_dso: str, expected_period: int = PERIOD) -> list[dict[str, Any]]:
    """Parse perf's physical no-inline stream, including every frame."""
    from decimal import Decimal, InvalidOperation

    rows: list[dict[str, Any]] = []
    current: dict[str, Any] | None = None

    def finish() -> None:
        nonlocal current
        if current is not None:
            rows.append(current)
            current = None

    for line_number, raw_line in enumerate(path.read_text(encoding="utf-8", errors="strict").splitlines(), 1):
        line = raw_line.rstrip("\n")
        stripped = line.strip()
        if not stripped:
            finish()
            continue
        header = _TEXT_HEADER.match(line)
        if header is not None:
            finish()
            event = header.group("event").strip()
            if event != "instructions:u":
                _fail("corrected_event", "corrected stream contains a different event", line=line_number, event=event)
            period = _positive_int(int(header.group("period")), f"corrected[{len(rows)+1}].period")
            if period != expected_period:
                _fail("corrected_period", "corrected stream period differs from bound event", line=line_number, period=period, expected=expected_period)
            try:
                time_ns = int(Decimal(header.group("time")) * Decimal(1_000_000_000))
            except (InvalidOperation, ValueError) as exc:
                _fail("corrected_time", "corrected stream timestamp is invalid", line=line_number)
                raise exc
            current = {
                "index": len(rows) + 1, "time_ns": time_ns, "period": period,
                "pid": int(header.group("pid")), "tid": int(header.group("tid") or header.group("pid")),
                "cpu": int(header.group("cpu")), "frames": [],
            }
            continue
        if current is not None and line[:1].isspace():
            frame = _TEXT_FRAME.match(line)
            if frame is None:
                _fail("corrected_frame", "malformed corrected frame", line=line_number, text=line[:240])
            symbol = frame.group("symbol").strip()
            base, _, offset = symbol.partition("+")
            dso = frame.group("dso").strip()
            current["frames"].append({
                "ip": _parse_perf_ip_token(frame.group("ip"), f"corrected[{len(rows)+1}].frame.ip"),
                "symbol": symbol, "symbol_base": base.strip(),
                "offset": offset or None, "dso": dso,
                "resolved": base.strip() not in {"[unknown]", "unknown", ""} and dso not in {"[unknown]", ""},
            })
            continue
        _fail("corrected_line", "unexpected corrected stream line", line=line_number, text=line[:240])
    finish()
    if not rows:
        _fail("corrected_samples_missing", "corrected stream has no samples", path=str(path))
    # Validate marker DSO on the independently parsed stream before consulting
    # the packet or attribution sidecar.
    for index, row in enumerate(rows, 1):
        for frame in row["frames"]:
            if frame["symbol_base"] == marker and frame["dso"] != marker_dso:
                _fail("marker_mismatch", "marker symbol resolved in unexpected DSO", sample=index, dso=frame["dso"], expected=marker_dso)
    return [_normalise_sample(row, i, str(path), marker=marker, marker_dso=marker_dso) for i, row in enumerate(rows, 1)]


def _parse_corrected(path: Path, obj: dict[str, Any] | None, *, marker: str, marker_dso: str) -> list[dict[str, Any]]:
    if obj is not None:
        return _parse_structured_rows(obj, "corrected", marker=marker, marker_dso=marker_dso)
    return _parse_corrected_text(path, marker=marker, marker_dso=marker_dso)


def _parse_raw(path: Path, obj: dict[str, Any] | None, *, marker: str, marker_dso: str) -> list[dict[str, Any]]:
    if obj is None:
        _fail("raw_sidecar_format", "raw sidecar must be structured JSON", path=str(path))
    if obj.get("schema") != "litchi.performance.0978.raw_stack.v1":
        _fail("raw_schema", "raw sidecar schema is not canonical", schema=obj.get("schema"))
    rows = _parse_structured_rows(obj, "raw", marker=marker, marker_dso=marker_dso)
    if obj.get("status") not in (None, "complete"):
        _fail("raw_status", "raw sidecar is not complete", status=obj.get("status"))
    declared = obj.get("sample_count")
    if declared is not None and declared != len(rows):
        _fail("raw_summary_mismatch", "raw sample_count differs from retained rows", declared=declared, actual=len(rows))
    declared_period = obj.get("total_period")
    actual_period = sum(row["period"] for row in rows)
    if declared_period is not None and declared_period != actual_period:
        _fail("raw_summary_mismatch", "raw total_period differs from retained rows", declared=declared_period, actual=actual_period)
    if obj.get("sample_conservation") is False:
        _fail("raw_conservation", "raw scanner did not prove sample conservation")
    return rows


def _load_sibling_module(name: str, filename: str) -> Any:
    path = Path(__file__).with_name(filename)
    spec = importlib.util.spec_from_file_location(name, path)
    if spec is None or spec.loader is None:
        _fail("helper_missing", f"cannot load bound helper {path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def _reparse_raw_scanner(
    raw_path: Path,
    raw_obj: Mapping[str, Any] | None,
    artifacts: Mapping[str, Mapping[str, Any]],
    *, marker: str,
    marker_dso: str,
    allow_synthetic: bool,
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]] | None:
    """Re-run the sealed variable-layout scanner over the bound raw dump.

    A packet produced by the real adapter must carry the scanner event
    contract and its dump binding.  Synthetic unit fixtures may omit those
    gigabyte-scale inputs, but the CLI never enables that escape hatch.
    """
    if not isinstance(raw_obj, Mapping) or not isinstance(raw_obj.get("event_contract"), Mapping):
        if allow_synthetic:
            return None
        _fail("raw_scanner_contract", "raw_stack lacks canonical event_contract for independent reparse")
    event_contract = dict(raw_obj["event_contract"])
    dump_binding = raw_obj.get("dump")
    dump_path = check_binding(dump_binding, "raw_stack.dump") if isinstance(dump_binding, Mapping) else None
    bound_dump = check_binding(artifacts["perf_mmap_dump"], "artifact.perf_mmap_dump")
    if dump_path is None or dump_path != bound_dump:
        _fail("raw_binding", "raw scanner dump is not the canonical perf_mmap_dump artifact")
    evlist_path = check_binding(artifacts["perf_evlist"], "artifact.perf_evlist")
    scanner = _load_sibling_module("litchi_raw_stack_0978_reducer", "raw_stack_0978.py")
    try:
        parsed_contract = scanner.parse_evlist_contract(
            evlist_path.read_text(encoding="utf-8", errors="strict"),
            expected_period=int(event_contract["period"]), expected_stack_cap=int(event_contract["stack_cap"]),
        )
        if parsed_contract != event_contract:
            _fail("raw_contract_mismatch", "raw_stack event contract differs from bound perf evlist")
        summary = scanner.scan_dump(
            dump_path,
            perf_data=check_binding(raw_obj["perf_data"], "raw_stack.perf_data") if isinstance(raw_obj.get("perf_data"), Mapping) else None,
            binary=check_binding(raw_obj["binary"], "raw_stack.binary") if isinstance(raw_obj.get("binary"), Mapping) else None,
            expected_pid=int(raw_obj["expected_pid"]), expected_cpu=int(raw_obj["expected_cpu"]),
            expected_period=int(event_contract["period"]), stack_cap=int(event_contract["stack_cap"]),
            event_contract=parsed_contract, retain_samples=True,
        )
    except Exception as exc:
        if isinstance(exc, ReducerFailure):
            raise
        _fail("raw_scanner_failure", f"sealed raw scanner failed: {exc}")
    if summary.get("status") != "complete" or summary.get("sample_conservation") is not True or summary.get("period_valid") is not True or summary.get("pid_tid_cpu_valid") is not True or summary.get("time_valid") is not True or summary.get("bounds_valid") is not True or summary.get("physical_identity_valid") is not True or summary.get("violations"):
        _fail("raw_scanner_failure", "sealed raw scanner did not prove complete sample conservation")
    if isinstance(raw_obj.get("sample_count"), int) and raw_obj["sample_count"] != summary["sample_count"]:
        _fail("raw_summary_mismatch", "raw_stack sample count differs from independent scanner", sidecar=raw_obj["sample_count"], scanner=summary["sample_count"])
    if isinstance(raw_obj.get("total_period"), int) and raw_obj["total_period"] != sum(int(row["period"]) for row in summary["sample_ledger"]):
        _fail("raw_summary_mismatch", "raw_stack period differs from independent scanner")
    census_keys = (
        "raw_event_count", "physical_location_count", "raw_sample_record_count",
        "raw_unconsumed_count", "sample_line_count", "raw_record_type_counts",
        "raw_record_size_counts", "mode_counts", "mode_missing_unwindability_counts",
        "regs_abi_counts", "stack_shape_counts",
    )
    bound_census = [key for key in census_keys if key in raw_obj]
    if not bound_census:
        _fail("raw_summary_mismatch", "raw_stack does not retain the raw census required for independent comparison")
    for key in bound_census:
        if raw_obj.get(key) != summary.get(key):
            _fail("raw_summary_mismatch", f"raw_stack {key} differs from independent scanner", sidecar=raw_obj.get(key), scanner=summary.get(key))
    original_rows = [dict(row) for row in summary["sample_ledger"]]
    normalized_rows = _parse_structured_rows({"sample_ledger": original_rows}, "raw scanner", marker=marker, marker_dso=marker_dso)
    return normalized_rows, original_rows


def _compare_built_lifecycle_rows(
    built_rows: Sequence[Mapping[str, Any]],
    declared_rows: Sequence[Mapping[str, Any]],
) -> None:
    """Compare every independent lifecycle classification with its sidecar.

    ``build_ledger`` keeps raw identity under ``raw_identity`` while the
    attribution sidecar keeps it at the row top level.  Compare both the
    identity and every disjoint classification field; accepting only the
    census would allow a phase or map-owner substitution to pass.
    """
    if len(built_rows) != len(declared_rows):
        _fail("lifecycle_attribution_mismatch", "built and declared lifecycle row counts differ", built=len(built_rows), declared=len(declared_rows))
    fields = ("phase", "mode", "marker", "physical_owner", "target_source_eligible", "unknown", "unknown_reason")
    identity = ("index", "time_ns", "pid", "tid", "cpu", "period")
    for built, declared in zip(built_rows, declared_rows):
        raw_identity = built.get("raw_identity")
        if not isinstance(raw_identity, Mapping):
            _fail("lifecycle_attribution_mismatch", "built lifecycle row lacks raw identity", index=declared.get("index"))
        index = declared.get("index")
        for key in identity:
            built_value = raw_identity.get(key)
            if built_value != declared.get(key):
                _fail("lifecycle_attribution_mismatch", f"built/declared lifecycle identity differs at {index}.{key}", built=built_value, declared=declared.get(key))
        for key in fields:
            if built.get(key) != declared.get(key):
                _fail("lifecycle_attribution_mismatch", f"built/declared lifecycle attribution differs at {index}.{key}", built=built.get(key), declared=declared.get(key))


def _compare_built_lifecycle_ledger(
    built_ledger: Mapping[str, Any],
    lifecycle_obj: Mapping[str, Any],
    packet: Mapping[str, Any] | None,
) -> None:
    """Check independent ledger scalars/maps against every retained claim."""
    declared = lifecycle_obj.get("ledger") if isinstance(lifecycle_obj.get("ledger"), Mapping) else lifecycle_obj
    packet_attr = packet.get("attribution") if isinstance(packet, Mapping) else None
    packet_ledger = packet_attr if isinstance(packet_attr, Mapping) and "ledger" not in packet_attr else None
    if isinstance(packet_attr, Mapping) and isinstance(packet_attr.get("ledger"), Mapping):
        # Historical decoder packets nested the ledger.  Live packets use the
        # canonical direct object, but comparing this copy here keeps the
        # custody check explicit if a retained packet is inspected offline.
        packet_ledger = packet_attr["ledger"]
    for source_name, source in (("lifecycle", declared), ("packet", packet_ledger)):
        if not isinstance(source, Mapping):
            continue
        for key in ("D", "M", "R", "T", "U_target", "U_marker", "marker_sample_count"):
            if key in source and source.get(key) != built_ledger.get(key):
                _fail("lifecycle_claim_mismatch", f"independent ledger differs from {source_name}.{key}", built=built_ledger.get(key), declared=source.get(key))
        for key in ("phase_periods", "target_known_owner_period", "target_source_eligible_period", "marker_known_owner_period"):
            if key in source and source.get(key) != built_ledger.get(key):
                _fail("lifecycle_claim_mismatch", f"independent ledger differs from {source_name}.{key}", built=built_ledger.get(key), declared=source.get(key))
        conservation = source.get("conservation")
        if conservation is not None and conservation != built_ledger.get("conservation"):
            _fail("lifecycle_claim_mismatch", f"independent ledger differs from {source_name}.conservation")


def _validate_marker_map_linkage(
    lifecycle_module: Any,
    stream: Mapping[str, Any],
    attributions: Mapping[int, Mapping[str, Any]],
    *,
    expected_pid: int,
    target_comm: str,
    marker_dso: str,
    target_build_id: str,
) -> None:
    if not any(isinstance(value, Mapping) and isinstance(value.get("marker_frame"), Mapping) for value in attributions.values()):
        return
    records = stream.get("records")
    samples = stream.get("sample_records")
    if not isinstance(records, list) or not isinstance(samples, list):
        _fail("attribution_binding", "lifecycle stream lacks records for marker map linkage")
    comms = [row for row in records if row.get("kind") == "PERF_RECORD_COMM" and row.get("pid") == expected_pid and row.get("tid") == expected_pid and row.get("exec") is True and str(row.get("comm")) == target_comm]
    if len(comms) != 1:
        _fail("lifecycle_identity", "cannot establish target COMM for marker map linkage", count=len(comms), target_comm=target_comm)
    comm_sequence = int(comms[0]["record_sequence"])
    maps = [row for row in records if row.get("kind") == "PERF_RECORD_MMAP2" and row.get("pid") == expected_pid and row.get("tid") == expected_pid]
    for index, attribution in attributions.items():
        marker_frame = attribution.get("marker_frame")
        if not isinstance(marker_frame, Mapping):
            continue
        if index <= 0 or index > len(samples):
            _fail("attribution_binding", "marker attribution index has no lifecycle sample", index=index)
        ip = _address(marker_frame.get("ip"), f"attributions[{index}].marker_frame.ip")
        sequence = int(samples[index - 1].get("record_sequence", -1))
        active = lifecycle_module._interval_owner(maps, ip, record_sequence=sequence, comm_sequence=comm_sequence)
        if not isinstance(active, Mapping):
            _fail("attribution_binding", "marker frame IP is not covered by an active map", index=index, ip=hex(ip))
        if active.get("path") != marker_dso or str(active.get("build_id") or "").lower() != target_build_id.lower():
            _fail("attribution_binding", "marker frame active map differs from exact marker DSO/BuildID", index=index)
        if marker_frame.get("map_record_sequence") != active.get("record_sequence"):
            _fail("attribution_binding", "marker frame map sequence differs from active lifecycle map", index=index, declared=marker_frame.get("map_record_sequence"), active=active.get("record_sequence"))


def _compare_built_ledger_to_actual(built_ledger: Mapping[str, Any], actual: Mapping[str, Any]) -> None:
    for key in ("D", "M", "R", "T", "U_target", "U_marker", "marker_sample_count"):
        if built_ledger.get(key) != actual.get(key):
            _fail("lifecycle_claim_mismatch", f"independent ledger differs from reducer-derived {key}", built=built_ledger.get(key), actual=actual.get(key))
    for built_key, actual_key in (
        ("phase_periods", "phase_periods"),
        ("target_known_owner_period", "target_known_owner_period"),
        ("target_source_eligible_period", "target_source_owner_period"),
        ("marker_known_owner_period", "marker_known_owner_period"),
    ):
        if built_ledger.get(built_key) != actual.get(actual_key):
            _fail("lifecycle_claim_mismatch", f"independent ledger differs from reducer-derived {actual_key}", built=built_ledger.get(built_key), actual=actual.get(actual_key))


def _reparse_lifecycle_stream(
    raw_dump_path: Path,
    raw_obj: Mapping[str, Any] | None,
    lifecycle_obj: Mapping[str, Any] | None,
    lifecycle_rows: Sequence[Mapping[str, Any]],
    *,
    raw_scanner_rows: Sequence[Mapping[str, Any]] | None = None,
    attribution_obj: Mapping[str, Any] | None = None,
    packet: Mapping[str, Any] | None = None,
    plan: Mapping[str, Any] | None = None,
    marker: str | None = None,
    marker_dso: str | None = None,
    allow_synthetic: bool,
) -> dict[str, Any] | None:
    """Independently frame the bound lifecycle/raw event stream.

    The lifecycle author owns attribution and phase policy.  The reducer still
    reruns the parser over the exact bound dump and checks its sample/loss
    census before accepting the author-owned ledger.
    """
    if allow_synthetic:
        return None
    if not isinstance(raw_obj, Mapping) or not isinstance(lifecycle_obj, Mapping):
        _fail("lifecycle_contract", "live lifecycle reparse inputs are not structured")
    if raw_scanner_rows is None or not isinstance(attribution_obj, Mapping):
        _fail("lifecycle_contract", "live lifecycle reparse lacks raw scanner rows or attribution sidecar")
    expected_pid = raw_obj.get("expected_pid")
    expected_cpu = raw_obj.get("expected_cpu")
    if not isinstance(expected_pid, int) or not isinstance(expected_cpu, int):
        _fail("lifecycle_contract", "raw_stack lacks expected PID/CPU for lifecycle reparse")
    lifecycle = _load_sibling_module("litchi_lifecycle_ledger_0978_reducer", "lifecycle_ledger_0978.py")
    try:
        stream = lifecycle.parse_lifecycle_dump(raw_dump_path, expected_pid=expected_pid, expected_cpu=expected_cpu, require_raw_metadata=True)
    except Exception as exc:
        if isinstance(exc, ReducerFailure):
            raise
        _fail("lifecycle_parser_failure", f"sealed lifecycle parser failed: {exc}")
    if stream.get("status") != "complete" or stream.get("raw_metadata_crosscheck") is not True or stream.get("unique_record_offsets") is not True:
        _fail("lifecycle_parser_failure", "lifecycle parser did not prove complete raw metadata custody")
    if stream.get("raw_lost_sum") != 0 or any(value != 0 for value in stream.get("terminal_lost_values", [])):
        _fail("loss_gate", "independent lifecycle stream has lost records", raw=stream.get("raw_lost_sum"), terminal=stream.get("terminal_lost_values"))
    for section_name in ("lifecycle", "gates"):
        section = lifecycle_obj.get(section_name)
        if not isinstance(section, Mapping):
            continue
        if section.get("raw_lost_sum", 0) != 0 or section.get("terminal_lost_samples", 0) != 0:
            _fail("loss_gate", f"lifecycle sidecar {section_name} reports loss", section=section_name)
    if int(stream.get("sample_descriptor_count", -1)) != len(lifecycle_rows):
        _fail("lifecycle_summary_mismatch", "lifecycle sample census differs from attribution sidecar", parser=stream.get("sample_descriptor_count"), sidecar=len(lifecycle_rows))
    declared_stream = lifecycle_obj.get("stream") or lifecycle_obj.get("lifecycle_stream")
    if isinstance(declared_stream, Mapping):
        for key in ("raw_event_count", "sample_descriptor_count", "raw_lost_sum"):
            if key in declared_stream and declared_stream[key] != stream.get(key):
                _fail("lifecycle_summary_mismatch", f"lifecycle {key} differs from independent parser", declared=declared_stream[key], actual=stream.get(key))

    # The sidecar's lifecycle rows are not an authority for phase or map
    # ownership.  Reconstruct the exact attribution handoff that the decoder
    # supplied to build_ledger and make the sealed helper perform the join
    # again over the raw scanner rows and freshly framed stream.
    raw_attributions = attribution_obj.get("attributions")
    if not isinstance(raw_attributions, Mapping):
        _fail("attribution_contract", "live attribution sidecar lacks exact attributions mapping")
    attributions: dict[int, Mapping[str, Any]] = {}
    for key, value in raw_attributions.items():
        try:
            index = int(key)
        except (TypeError, ValueError):
            _fail("attribution_contract", "attribution index is not an integer", index=key)
        if index <= 0 or index > len(raw_scanner_rows):
            _fail("attribution_contract", "attribution index is out of range", index=index)
        if not isinstance(value, Mapping):
            _fail("attribution_contract", "attribution entry is not an object", index=index)
        attributions[index] = dict(value)

    # Every non-empty attribution claim must have the exact map/build-id
    # binding needed by build_ledger.  Unknown rows may legitimately be
    # omitted from the mapping, but a row that the lifecycle sidecar calls
    # owned, target-source eligible, or marker-bearing may not be silently
    # downgraded to unknown by this reducer.
    for row in lifecycle_rows:
        index = int(row["index"])
        claim = attributions.get(index)
        if (row.get("physical_owner") is not None or row.get("target_source_eligible") is True or row.get("marker") is True) and claim is None:
            _fail("attribution_contract", "lifecycle row has an attribution claim absent from exact mapping", index=index)
        if row.get("target_source_eligible") is not True and row.get("target_source_eligible") is not False:
            _fail("attribution_contract", "lifecycle target_source_eligible is missing/non-boolean", index=index)
        if not isinstance(row.get("phase"), str) or not isinstance(row.get("unknown"), bool):
            _fail("attribution_contract", "lifecycle phase/unknown fields are missing", index=index)

    # Prefer the plan's immutable binary identity.  The packet/sidecar may
    # repeat it, but those copies are checked against the plan below.
    binary = plan.get("binary") if isinstance(plan, Mapping) and isinstance(plan.get("binary"), Mapping) else {}
    target_path = str(binary.get("path") or "")
    target_build_id = str((plan.get("marker") or {}).get("build_id") if isinstance(plan, Mapping) and isinstance(plan.get("marker"), Mapping) else "").lower()
    target_comm = str((plan.get("workload") or {}).get("target_comm") if isinstance(plan, Mapping) and isinstance(plan.get("workload"), Mapping) else "")
    if not target_comm:
        target_comm = Path(target_path).name[:15]
    sidecar_target = lifecycle_obj.get("target")
    packet_target = packet.get("target") if isinstance(packet, Mapping) else None
    for source_name, value in (("lifecycle", sidecar_target), ("packet", packet_target)):
        if value is None:
            continue
        if not isinstance(value, Mapping):
            _fail("lifecycle_target", f"{source_name} target binding is not an object")
        if target_path and value.get("path") not in (None, target_path):
            _fail("lifecycle_target", f"{source_name} target path differs from plan", actual=value.get("path"), expected=target_path)
        if target_build_id and str(value.get("build_id") or "").lower() not in {"", target_build_id}:
            _fail("lifecycle_target", f"{source_name} target BuildID differs from plan", actual=value.get("build_id"), expected=target_build_id)
        if value.get("comm") not in (None, target_comm):
            _fail("lifecycle_target", f"{source_name} target COMM differs from plan", actual=value.get("comm"), expected=target_comm)
    if not target_path or not re.fullmatch(r"[0-9a-f]{8,64}", target_build_id):
        _fail("lifecycle_target", "plan lacks exact target path/BuildID for ledger replay")

    gate = packet.get("gate") if isinstance(packet, Mapping) else None
    terminal_lost = None
    if isinstance(gate, Mapping):
        terminal_lost = gate.get("terminal_lost_samples")
    for value in (
        lifecycle_obj.get("terminal_lost_samples"),
        (lifecycle_obj.get("lifecycle") or {}).get("terminal_lost_samples") if isinstance(lifecycle_obj.get("lifecycle"), Mapping) else None,
    ):
        if value is not None:
            if type(value) is not int:
                _fail("loss_gate", "lifecycle terminal loss is not an integer", value=value)
            if terminal_lost is None:
                terminal_lost = value
            elif terminal_lost != value:
                _fail("loss_gate", "packet/lifecycle terminal loss differs", packet=terminal_lost, lifecycle=value)
    if terminal_lost is None:
        _fail("loss_gate", "terminal PERF_RECORD_LOST_SAMPLES value is not bound")

    _validate_marker_map_linkage(
        lifecycle,
        stream,
        attributions,
        expected_pid=expected_pid,
        target_comm=target_comm,
        marker_dso=str(marker_dso or target_path),
        target_build_id=target_build_id,
    )

    try:
        built = lifecycle.build_ledger(
            stream,
            raw_scanner_rows,
            expected_pid=expected_pid,
            expected_cpu=expected_cpu,
            expected_period=PERIOD,
            target_path=target_path,
            target_build_id=target_build_id,
            target_comm=target_comm,
            terminal_lost_samples=terminal_lost,
            marker=marker,
            marker_dso=marker_dso,
            attributions=attributions,
        )
    except Exception as exc:
        if isinstance(exc, ReducerFailure):
            raise
        _fail("lifecycle_ledger_failure", f"independent build_ledger rejected bound streams: {exc}")
    if not isinstance(built, Mapping) or built.get("status") != "complete":
        _fail("lifecycle_ledger_failure", "independent build_ledger did not return a complete ledger")
    built_rows = built.get("samples")
    built_ledger = built.get("ledger")
    if not isinstance(built_rows, list) or not isinstance(built_ledger, Mapping):
        _fail("lifecycle_ledger_failure", "independent build_ledger omitted samples or ledger")
    _compare_built_lifecycle_rows(built_rows, lifecycle_rows)
    _compare_built_lifecycle_ledger(built_ledger, lifecycle_obj, packet)
    return dict(built)


def _parse_lifecycle(path: Path, obj: dict[str, Any] | None, *, marker: str, marker_dso: str) -> tuple[list[dict[str, Any]], dict[str, Any]]:
    if obj is None:
        _fail("lifecycle_sidecar_format", "lifecycle sidecar must be structured JSON", path=str(path))
    if obj.get("schema") != "litchi.performance.0978.lifecycle_ledger.v1":
        _fail("lifecycle_schema", "lifecycle sidecar schema is not canonical", schema=obj.get("schema"))
    rows = _parse_structured_rows(obj, "lifecycle", marker=marker, marker_dso=marker_dso, allow_summary_claims=True, allow_unframed_claims=True)
    ledger = obj.get("ledger") if isinstance(obj.get("ledger"), Mapping) else obj
    if not isinstance(ledger, Mapping):
        _fail("lifecycle_shape", "lifecycle ledger is not an object")
    for key in ("D", "M", "R", "T", "U_target", "U_marker"):
        if key in ledger:
            _nonnegative_int(ledger[key], f"lifecycle.{key}")
    return rows, dict(ledger)


def _require_live_lifecycle_contract(
    lifecycle_obj: Mapping[str, Any] | None,
    lifecycle_rows: Sequence[Mapping[str, Any]],
    ledger: Mapping[str, Any],
) -> None:
    if not isinstance(lifecycle_obj, Mapping):
        _fail("lifecycle_contract", "live lifecycle sidecar is missing")
    required = ("D", "M", "R", "T", "U_target", "U_marker", "marker_sample_count")
    missing = [key for key in required if key not in ledger]
    if missing:
        _fail("lifecycle_contract", "live lifecycle ledger omits core totals", missing=missing)
    for key in required:
        _nonnegative_int(ledger[key], f"lifecycle.{key}")
    required_maps = ("target_known_owner_period", "target_source_eligible_period", "marker_known_owner_period")
    for key in required_maps:
        if key not in ledger:
            _fail("lifecycle_contract", "live lifecycle ledger omits owner map", field=key)
        _map_numbers(ledger[key], f"lifecycle.{key}")
    conservation = ledger.get("conservation")
    if not isinstance(conservation, Mapping) or any(value is not True for value in conservation.values()):
        _fail("lifecycle_contract", "live lifecycle ledger conservation gate is missing or failed")
    for row in lifecycle_rows:
        index = row.get("index")
        if not isinstance(row.get("phase"), str) or not isinstance(row.get("mode"), str):
            _fail("lifecycle_contract", "live lifecycle sample omits phase or mode", index=index)
        if not isinstance(row.get("marker"), bool) or not isinstance(row.get("unknown"), bool):
            _fail("lifecycle_contract", "live lifecycle sample omits marker or unknown classification", index=index)
        if row.get("target_source_eligible") not in (True, False):
            _fail("lifecycle_contract", "live lifecycle sample omits target-source eligibility", index=index)


def _compare_identities(left: Sequence[Mapping[str, Any]], right: Sequence[Mapping[str, Any]], left_name: str, right_name: str) -> None:
    if len(left) != len(right):
        _fail("sample_conservation", f"{left_name}/{right_name} sample counts differ", left=len(left), right=len(right))
    for a, b in zip(left, right):
        for key in ("index", "period", "time_ns", "pid", "tid", "cpu"):
            # A lifecycle/raw sidecar may omit optional pid/tid/cpu only when
            # the packet contract says so.  If both retain a value, it must be
            # identical; period/time/index are always mandatory.
            if key in ("index", "period", "time_ns") or (a.get(key) is not None and b.get(key) is not None):
                left_value, right_value = a.get(key), b.get(key)
                # perf script prints timestamps at microsecond precision while
                # the raw sample payload retains nanoseconds.  The retained
                # 0977 join contract compares those two sources at 1us; the
                # lifecycle/raw build_ledger join below remains exact.
                if key == "time_ns" and ("corrected" in left_name or "corrected" in right_name):
                    left_value, right_value = int(left_value) // 1000, int(right_value) // 1000
                if left_value != right_value:
                    _fail("sample_identity", f"{left_name}/{right_name} differ at sample {a.get('index')} field {key}", left=a.get(key), right=b.get(key))


def _frame_signature(frame: Mapping[str, Any]) -> tuple[Any, ...]:
    ip = frame.get("ip")
    if isinstance(ip, str):
        try:
            ip = int(ip, 0)
        except ValueError:
            try:
                ip = int(ip, 16)
            except ValueError:
                pass
    return (
        ip,
        _frame_dso(frame),
        _frame_symbol(frame),
        frame.get("symbol_base"),
        frame.get("offset"),
        _frame_resolved(frame),
    )


def _compare_attribution_rows(left: Sequence[Mapping[str, Any]], right: Sequence[Mapping[str, Any]], left_name: str, right_name: str) -> None:
    _compare_identities(left, right, left_name, right_name)
    for a, b in zip(left, right):
        index = a.get("index")
        for key in ("marker", "physical_owner", "inclusive_owners"):
            if a.get(key) != b.get(key):
                _fail("attribution_mismatch", f"{left_name}/{right_name} differ at sample {index} field {key}", left=a.get(key), right=b.get(key))
        # The corrected stream is the independent frame identity source.  Do
        # not accept equal owner totals when the sidecar swapped an IP,
        # symbol, DSO, resolution bit, or frame order.
        if "frames" in a and "frames" in b:
            left_frames = [_frame_signature(frame) for frame in a["frames"]]
            right_frames = [_frame_signature(frame) for frame in b["frames"]]
            if left_frames != right_frames:
                _fail("attribution_mismatch", f"{left_name}/{right_name} frame stacks differ at sample {index}", left=left_frames, right=right_frames)


def _compare_lifecycle_rows(left: Sequence[Mapping[str, Any]], right: Sequence[Mapping[str, Any]], left_name: str, right_name: str) -> None:
    _compare_attribution_rows(left, right, left_name, right_name)
    for a, b in zip(left, right):
        index = a.get("index")
        for key in ("phase", "mode", "target_source_eligible", "unknown", "unknown_reason"):
            if a.get(key) != b.get(key):
                _fail("lifecycle_attribution_mismatch", f"{left_name}/{right_name} differ at sample {index} field {key}", left=a.get(key), right=b.get(key))


def _address(value: Any, label: str) -> int:
    if type(value) is int and value >= 0:
        return value
    if isinstance(value, str):
        try:
            return int(value, 0)
        except ValueError:
            try:
                return int(value, 16)
            except ValueError:
                pass
    _fail("attribution_binding", f"{label} is not a valid address", value=value)


def _validate_exact_attribution_linkage(
    attribution_obj: Mapping[str, Any],
    corrected: Sequence[Mapping[str, Any]],
    *,
    marker: str,
    marker_dso: str,
    target_build_id: str,
    allow_synthetic: bool,
) -> None:
    """Bind each owner/marker claim to the independently parsed frame stack."""
    if allow_synthetic:
        return
    raw_attributions = attribution_obj.get("attributions")
    if not isinstance(raw_attributions, Mapping):
        _fail("attribution_contract", "live attribution sidecar lacks attributions mapping")
    rows = {int(row["index"]): row for row in corrected}
    for key, value in raw_attributions.items():
        try:
            index = int(key)
        except (TypeError, ValueError):
            _fail("attribution_binding", "attribution index is not an integer", index=key)
        row = rows.get(index)
        if row is None or not isinstance(value, Mapping):
            _fail("attribution_binding", "attribution entry has no corrected sample", index=index)
        frames = row.get("frames", [])
        if not isinstance(frames, list):
            _fail("attribution_binding", "corrected sample frames are not a list", index=index)
        leaf = value.get("physical_leaf")
        if isinstance(leaf, Mapping):
            if not frames or not _frame_resolved(frames[0]):
                _fail("attribution_binding", "physical leaf is claimed without a resolved corrected frame-0", index=index)
            frame = frames[0]
            frame_ip = _address(frame.get("ip"), f"corrected[{index}].frame0.ip")
            leaf_ip = _address(leaf.get("ip"), f"attributions[{index}].physical_leaf.ip")
            if leaf_ip != frame_ip:
                _fail("attribution_binding", "physical leaf IP differs from corrected frame-0", index=index, leaf=leaf_ip, frame=frame_ip)
            if leaf.get("dso_path") != _frame_dso(frame):
                _fail("attribution_binding", "physical leaf DSO differs from corrected frame-0", index=index)
            if leaf.get("symbol") != (_frame_symbol(frame) or ""):
                _fail("attribution_binding", "physical leaf symbol differs from corrected frame-0", index=index, leaf=leaf.get("symbol"), frame=_frame_symbol(frame))
            if type(leaf.get("map_record_sequence")) is not int or leaf.get("map_record_sequence") <= 0:
                _fail("attribution_binding", "physical leaf map sequence is missing", index=index)
            expected_owner = f"{leaf.get('dso_path')}::{leaf.get('symbol')}"
            if row.get("physical_owner") != expected_owner:
                _fail("attribution_binding", "physical leaf owner differs from corrected frame-0 owner", index=index, claim=expected_owner, corrected=row.get("physical_owner"))
            expected_target = leaf.get("dso_path") == marker_dso and str(leaf.get("build_id") or "").lower() == target_build_id.lower()
            if value.get("target_source_eligible") is not expected_target:
                _fail("attribution_binding", "physical leaf target-source eligibility disagrees with exact identity", index=index)
        elif value.get("physical_unknown") is not True and not isinstance(value.get("marker_frame"), Mapping):
            _fail("attribution_binding", "attribution entry lacks physical leaf or explicit unknown marker binding", index=index)
        marker_frame = value.get("marker_frame")
        if marker_frame is not None:
            if not isinstance(marker_frame, Mapping) or marker_frame.get("present") is not True:
                _fail("attribution_binding", "marker_frame is not an explicit present object", index=index)
            marker_candidates = [frame for frame in frames if _frame_symbol(frame) == marker and _frame_dso(frame) == marker_dso]
            if not marker_candidates:
                _fail("attribution_binding", "marker_frame has no matching corrected frame", index=index)
            frame = marker_candidates[0]
            if marker_frame.get("symbol") != marker or marker_frame.get("dso_path") != marker_dso or str(marker_frame.get("build_id") or "").lower() != target_build_id.lower():
                _fail("attribution_binding", "marker_frame identity differs from plan marker", index=index)
            if marker_frame.get("resolved") is not _frame_resolved(frame):
                _fail("attribution_binding", "marker_frame resolution differs from corrected frame", index=index)
            if "ip" not in marker_frame:
                _fail("attribution_binding", "marker_frame omits exact IP linkage", index=index)
            if _address(marker_frame.get("ip"), f"attributions[{index}].marker_frame.ip") != _address(frame.get("ip"), f"corrected[{index}].marker.ip"):
                _fail("attribution_binding", "marker_frame IP differs from corrected marker frame", index=index)
    for index, row in rows.items():
        if row.get("marker") is True and index not in {int(key) for key in raw_attributions}:
            _fail("attribution_binding", "corrected marker sample lacks exact attribution entry", index=index)


def _claim_object(packet: Mapping[str, Any]) -> Mapping[str, Any]:
    attribution = packet.get("attribution")
    if not isinstance(attribution, Mapping):
        _fail("packet_claim_missing", f"{packet.get('packet_id')}: attribution ledger is missing")
    if "ledger" in attribution:
        _fail("packet_claim_shape", f"{packet.get('packet_id')}: attribution must be the lifecycle ledger object, not a nested alias")
    combined = dict(attribution)
    # The inherited corrected decoder summary remains an independent bounded
    # claim source.  Merge it only when each duplicate field is identical.
    perf = packet.get("perf")
    corrected = perf.get("corrected") if isinstance(perf, Mapping) else None
    if corrected is not None:
        if not isinstance(corrected, Mapping):
            _fail("packet_claim_missing", f"{packet.get('packet_id')}: perf.corrected is not an object")
        for key, value in corrected.items():
            if key in combined and combined[key] != value:
                _fail("packet_claim_mismatch", f"{packet.get('packet_id')}: perf.corrected and attribution differ", field=key)
            combined.setdefault(key, value)
    return combined


def _claim(claims: Mapping[str, Any], names: Sequence[str], label: str) -> Any:
    value = _first_mapping(claims, names)
    if value is None:
        _fail("packet_claim_missing", f"{label} is missing", aliases=list(names))
    return value


def _map_numbers(value: Any, label: str) -> dict[str, int]:
    if not isinstance(value, Mapping):
        _fail("owner_map_shape", f"{label} is not an object")
    result: dict[str, int] = {}
    for key, raw in value.items():
        if not isinstance(key, str) or not key:
            _fail("owner_map_shape", f"{label} has invalid owner key")
        result[key] = _nonnegative_int(raw, f"{label}[{key}]")
    return result


def _sum_map(value: Mapping[str, int]) -> int:
    return sum(value.values())


def _derive_child(corrected: Sequence[Mapping[str, Any]], raw: Sequence[Mapping[str, Any]], lifecycle: Sequence[Mapping[str, Any]], ledger_declared: Mapping[str, Any], *, marker_dso: str) -> dict[str, Any]:
    _compare_identities(corrected, raw, "corrected", "raw")
    _compare_identities(corrected, lifecycle, "corrected", "lifecycle")
    d = sum(int(row["period"]) for row in raw)
    m = 0
    marker_count = 0
    physical: Counter[str] = Counter()
    inclusive: Counter[str] = Counter()
    physical_unknown = 0
    marker_unknown_samples = 0
    target_owner: Counter[str] = Counter()
    target_source_owner: Counter[str] = Counter()
    target_known = 0
    target_unknown = 0
    phase_periods: Counter[str] = Counter()
    phase_counts: Counter[str] = Counter()
    for index, (c_row, l_row) in enumerate(zip(corrected, lifecycle), 1):
        if bool(c_row["marker"]) != bool(l_row["marker"]):
            _fail("marker_mismatch", f"corrected/lifecycle marker disagreement at sample {index}")
        marker = bool(c_row["marker"])
        phase = l_row.get("phase") or "unknown"
        phase_periods[phase] += int(c_row["period"])
        phase_counts[phase] += 1
        if marker:
            marker_count += 1
            m += int(c_row["period"])
            if c_row["physical_owner"] is None or l_row.get("unknown") is True:
                physical_unknown += int(c_row["period"])
                marker_unknown_samples += 1
            else:
                physical[c_row["physical_owner"]] += int(c_row["period"])
            for owner in c_row["inclusive_owners"]:
                inclusive[owner] += int(c_row["period"])
        if phase == "target_execution" or l_row.get("target_window") is True:
            eligible = l_row.get("target_source_eligible")
            owner = l_row.get("physical_owner")
            if owner is not None and l_row.get("unknown") is not True:
                target_owner[owner] += int(c_row["period"])
                target_known += int(c_row["period"])
            else:
                target_unknown += int(c_row["period"])
            if eligible is True and owner is not None and l_row.get("unknown") is not True:
                target_source_owner[owner] += int(c_row["period"])
        # The lifecycle ledger deliberately scopes physical ownership to the
        # target/marker phase.  A pre-exec corrected frame may have a mapped
        # startup leaf while the lifecycle handoff keeps that phase unknown;
        # that is not a substitution.  Once a sample contributes to target or
        # marker attribution, every owner claim must agree exactly.
        lifecycle_owner = l_row.get("physical_owner")
        corrected_owner = c_row.get("physical_owner")
        if lifecycle_owner is not None and lifecycle_owner != corrected_owner:
            _fail("owner_mismatch", f"corrected/lifecycle physical owner disagreement at sample {index}")
        if (phase == "target_execution" or marker) and corrected_owner is not None and lifecycle_owner != corrected_owner:
            _fail("owner_mismatch", f"target/marker physical owner is missing from lifecycle at sample {index}")
    r = d - m
    if r < 0:
        _fail("period_conservation", "marker period exceeds D", D=d, M=m)
    t = target_known + target_unknown
    u_marker = physical_unknown
    if _sum_map(physical) + u_marker != m:
        _fail("period_conservation", "marker physical-leaf buckets do not conserve M")
    if t != phase_periods.get("target_execution", 0) and not any(row.get("target_window") is True for row in lifecycle):
        _fail("period_conservation", "target phase periods do not conserve T", T=t, phase=phase_periods.get("target_execution", 0))
    # WIP lifecycle ledger must be independently consistent with the rows.
    expected_ledger = {"D": d, "M": m, "R": r, "T": t, "U_target": target_unknown, "U_marker": u_marker}
    for key, actual in expected_ledger.items():
        if key in ledger_declared and ledger_declared[key] != actual:
            _fail("lifecycle_claim_mismatch", f"lifecycle {key} differs from reparsed rows", declared=ledger_declared[key], actual=actual)
    if "target_known_owner_period" in ledger_declared:
        declared_map = _map_numbers(ledger_declared["target_known_owner_period"], "lifecycle.target_known_owner_period")
        if declared_map != dict(target_owner):
            _fail("lifecycle_claim_mismatch", "lifecycle target owner map differs from rows")
    if "target_source_eligible_period" in ledger_declared:
        declared_map = _map_numbers(ledger_declared["target_source_eligible_period"], "lifecycle.target_source_eligible_period")
        if declared_map != dict(target_source_owner):
            _fail("lifecycle_claim_mismatch", "lifecycle target-source owner map differs from rows")
    if "marker_known_owner_period" in ledger_declared:
        declared_map = _map_numbers(ledger_declared["marker_known_owner_period"], "lifecycle.marker_known_owner_period")
        if declared_map != dict(physical):
            _fail("lifecycle_claim_mismatch", "lifecycle marker owner map differs from rows")
    return {
        "D": d, "M": m, "R": r, "T": t, "Tknown": target_known,
        "U_target": target_unknown, "U_marker": u_marker,
        "sample_count": len(raw), "marker_sample_count": marker_count,
        "physical_leaf_owner_period": dict(sorted(physical.items())),
        "marker_known_owner_period": dict(sorted(physical.items())),
        "physical_leaf_unknown_period": physical_unknown,
        "physical_leaf_unknown_samples": marker_unknown_samples,
        "inclusive_owner_period": dict(sorted(inclusive.items())),
        "target_known_owner_period": dict(sorted(target_owner.items())),
        "target_source_owner_period": dict(sorted(target_source_owner.items())),
        "phase_periods": dict(sorted(phase_periods.items())),
        "phase_counts": dict(sorted(phase_counts.items())),
    }


def _compare_packet_claims(packet: Mapping[str, Any], claims: Mapping[str, Any], actual: Mapping[str, Any]) -> None:
    pid = packet.get("packet_id")
    required = {
        "D": "D", "M": "M", "R": "R", "T": "T",
        "U_target": "U_target", "U_marker": "U_marker",
        "sample_count": "sample_count", "marker_sample_count": "marker_sample_count",
    }
    for key, name in required.items():
        value = _claim(claims, (name,), f"{pid}.{key}")
        expected = actual[key]
        if value != expected:
            _fail("packet_claim_mismatch", f"{pid}.{key} differs from reparsed sidecars", claimed=value, actual=expected)
    tknown = claims.get("Tknown")
    if tknown is None:
        owner_map = _map_numbers(_claim(claims, ("target_known_owner_period",), f"{pid}.target_known_owner_period"), f"{pid}.target_known_owner_period")
        tknown = sum(owner_map.values())
    if tknown != actual["Tknown"]:
        _fail("packet_claim_mismatch", f"{pid}.Tknown differs from reparsed sidecars", claimed=tknown, actual=actual["Tknown"])
    map_fields = {
        "physical_leaf_owner_period": "physical_leaf_owner_period",
        # The inherited corrected summary calls this view
        # ``inclusive_marker_ancestor_period``; 0978 attribution may expose
        # the shorter owner name only in its separate ledger.
        "inclusive_owner_period": ("inclusive_marker_ancestor_period", "inclusive_owner_period"),
        "marker_known_owner_period": "marker_known_owner_period",
        "target_known_owner_period": "target_known_owner_period",
        "target_source_owner_period": "target_source_eligible_period",
    }
    for key, names in map_fields.items():
        names_tuple = (names,) if isinstance(names, str) else names
        value = _map_numbers(_claim(claims, names_tuple, f"{pid}.{key}"), f"{pid}.{key}")
        expected_value = actual["physical_leaf_owner_period"] if key == "marker_known_owner_period" else actual[key]
        if value != expected_value:
            _fail("packet_claim_mismatch", f"{pid}.{key} differs from reparsed sidecars", claimed=value, actual=actual[key])
    for key in ("physical_leaf_unknown_period", "physical_leaf_unknown_samples"):
        value = _claim(claims, (key,), f"{pid}.{key}")
        if value != actual[key]:
            _fail("packet_claim_mismatch", f"{pid}.{key} differs from reparsed sidecars", claimed=value, actual=actual[key])


def _percentile(values: Sequence[float], fraction: float) -> float:
    if not values:
        _fail("empty_metric", "cannot calculate percentile of empty vector")
    ordered = sorted(float(v) for v in values)
    position = fraction * (len(ordered) - 1)
    low, high = math.floor(position), math.ceil(position)
    if low == high:
        return ordered[low]
    return ordered[low] + (ordered[high] - ordered[low]) * (position - low)


def _rank(values: Sequence[float]) -> list[float]:
    pairs = sorted((value, index) for index, value in enumerate(values))
    result = [0.0] * len(values)
    cursor = 0
    while cursor < len(pairs):
        end = cursor + 1
        while end < len(pairs) and pairs[end][0] == pairs[cursor][0]:
            end += 1
        rank = (cursor + 1 + end) / 2.0
        for _, index in pairs[cursor:end]:
            result[index] = rank
        cursor = end
    return result


def _correlation(left: Sequence[float], right: Sequence[float]) -> dict[str, float | None]:
    if len(left) != len(right) or not left:
        _fail("vector_shape", "correlation vectors are not paired")
    def corr(a: Sequence[float], b: Sequence[float]) -> float | None:
        ma, mb = sum(a) / len(a), sum(b) / len(b)
        da = sum((x - ma) ** 2 for x in a)
        db = sum((y - mb) ** 2 for y in b)
        if da == 0 or db == 0:
            return None
        return sum((x - ma) * (y - mb) for x, y in zip(a, b)) / math.sqrt(da * db)
    return {"pearson": corr(left, right), "spearman": corr(_rank(left), _rank(right))}


def _distribution(values: Sequence[float | int]) -> dict[str, Any]:
    numbers = [float(value) for value in values]
    mean = sum(numbers) / len(numbers)
    variance = sum((value - mean) ** 2 for value in numbers) / len(numbers)
    return {
        "count": len(numbers), "min": min(numbers), "median": _percentile(numbers, 0.5), "max": max(numbers),
        "mean": mean, "stddev_population": math.sqrt(variance),
        "cv_population": math.sqrt(variance) / mean if mean else None,
    }


def _child_timing(packet: Mapping[str, Any], artifacts: Mapping[str, Mapping[str, Any]], *, expected_count: int | None = None) -> list[float]:
    timing = packet.get("timing")
    if isinstance(timing, Mapping):
        values = timing.get("elapsed_ns", timing.get("samples"))
    else:
        values = None
    if values is None and isinstance(packet.get("report"), Mapping):
        values = packet["report"].get("elapsed_ns")
    if values is None and "report" in artifacts:
        report = load_json(check_binding(artifacts["report"], f"{packet.get('packet_id')}.report"), "report")
        values = report.get("elapsed_ns")
    if not isinstance(values, list) or not values:
        _fail("timing_missing", f"{packet.get('packet_id')}: no elapsed timing vector")
    result = [_finite_number(value, f"{packet.get('packet_id')}.elapsed_ns") for value in values]
    if any(value < 0 for value in result):
        _fail("timing_value", f"{packet.get('packet_id')}: elapsed timing is negative")
    if expected_count is not None and len(result) != expected_count:
        _fail("timing_count", f"{packet.get('packet_id')}: report timing count differs from plan", actual=len(result), expected=expected_count)
    return result


_RETAINED_0977_REDUCER: Any | None = None


def _load_retained_0977_reducer() -> Any:
    """Load the immutable 0977 custody/oracle checks without patching them."""
    global _RETAINED_0977_REDUCER
    if _RETAINED_0977_REDUCER is not None:
        return _RETAINED_0977_REDUCER
    path = Path("/home/zhuhe/litchi-performance-0977/tools/reduce_0977.py")
    if not path.is_file():
        _fail("custody_helper_missing", "retained 0977 reducer is missing", path=str(path))
    spec = importlib.util.spec_from_file_location("litchi_reduce_0977_custody_for_0978", path)
    if spec is None or spec.loader is None:
        _fail("custody_helper_missing", "cannot load retained 0977 reducer", path=str(path))
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    try:
        spec.loader.exec_module(module)
    except Exception as exc:
        _fail("custody_helper_missing", f"retained 0977 reducer import failed: {exc}")
    _RETAINED_0977_REDUCER = module
    return module


def _check_packet_plan_binding(packet: Mapping[str, Any], plan_path: Path, plan: Mapping[str, Any]) -> None:
    value = packet.get("plan")
    if not isinstance(value, Mapping):
        _fail("plan_binding", f"{packet.get('packet_id')}: packet plan binding is missing")
    if value.get("path") != str(plan_path.absolute()):
        _fail("plan_binding", f"{packet.get('packet_id')}: packet plan path differs")
    if value.get("sha256") != sha256(plan_path):
        _fail("plan_binding", f"{packet.get('packet_id')}: packet plan digest differs")
    if value.get("schema") != plan.get("schema"):
        _fail("plan_binding", f"{packet.get('packet_id')}: packet plan schema differs")


def _check_packet_output_boundary(packet_path: Path, packet: Mapping[str, Any], plan: Mapping[str, Any]) -> None:
    outputs = plan.get("outputs")
    if not isinstance(outputs, Mapping) or not isinstance(outputs.get("capture_root"), str):
        _fail("plan_contract", "plan output capture root is missing")
    expected_root = Path(outputs["capture_root"]).absolute()
    packet_id = packet.get("packet_id")
    expected_dir = expected_root / str(packet_id)
    if packet_path.parent != expected_dir:
        _fail("packet_custody", f"{packet_id}: packet is outside plan capture child directory", actual=str(packet_path.parent), expected=str(expected_dir))


def _check_child_artifact_boundaries(packet_path: Path, artifacts: Mapping[str, Mapping[str, Any]], packet_id: str) -> None:
    child_dir = packet_path.parent
    for name, row in artifacts.items():
        path = check_binding(row, f"{packet_id}.artifact.{name}")
        if path.parent != child_dir:
            _fail("packet_custody", f"{packet_id}: child artifact escapes packet directory", artifact=name, path=str(path), expected_parent=str(child_dir))


def _compare_legacy_decoder_summary(base: Any, packet: Mapping[str, Any], parsed: Mapping[str, Any], losses: Mapping[str, int]) -> Mapping[str, Any]:
    """Apply the unchanged 0977 summary invariants without old owner rules.

    The retained checker still owns the strict corrected perf-script parser
    and all scalar/sample-count invariants.  Its physical-leaf and unknown
    fields used a different owner contract, so they are projected from the
    independently parsed stream and validated later by the 0978 lifecycle
    ledger.  New 0978 fields are allowlisted here and remain claim-checked
    by ``_compare_packet_claims``.
    """
    perf = packet.get("perf")
    if not isinstance(perf, Mapping) or set(perf) != {"corrected"}:
        _fail("decoder_summary_contract", f"{packet.get('packet_id')}: packet must contain corrected-only perf summary")
    summary = perf.get("corrected")
    if not isinstance(summary, Mapping) or "samples" in summary:
        _fail("decoder_summary_contract", f"{packet.get('packet_id')}: decoder summary is missing/bloated with sample rows")
    required = set(getattr(base, "REQUIRED_SUMMARY_KEYS", ()))
    allowed = set(getattr(base, "ALLOWED_SUMMARY_EXTRAS", ()))
    if not required:
        _fail("decoder_summary_contract", "retained checker exposes no required summary contract")
    unsupported = sorted(set(summary) - required - allowed - set(SUMMARY_0978_KEYS))
    if unsupported:
        _fail("decoder_summary_contract", f"{packet.get('packet_id')}: unexpected 0978 decoder summary keys", keys=unsupported)
    # Current 0978 emits these unchanged scalar invariants.  Require them in
    # the packet rather than silently filling them from the old parser.
    emitted = {
        "sample_count", "total_period", "marker_sample_count", "marker_period",
        "marker_period_share", "lost_events", "unresolved_frames",
        "explicit_stack_truncated_count", "empty_callchain_count",
    }
    missing_emitted = sorted(key for key in emitted if key not in summary)
    if missing_emitted:
        _fail("decoder_summary_contract", f"{packet.get('packet_id')}: unchanged decoder fields missing", keys=missing_emitted)
    projected: dict[str, Any] = {}
    for key in required:
        if key in LEGACY_OWNER_SUMMARY_KEYS:
            # Old and new owner/unknown definitions are intentionally
            # different.  The new ledger and packet claims check these.
            if key not in parsed:
                _fail("decoder_summary_contract", f"retained parser omitted projected field {key}")
            projected[key] = parsed[key]
        elif key in summary:
            projected[key] = summary[key]
        elif key in parsed:
            # ``marker`` and parser-only scope labels are bound by the plan
            # and strict parser; 0978 does not duplicate every old label.
            projected[key] = parsed[key]
        else:
            _fail("decoder_summary_contract", f"{packet.get('packet_id')}: summary/parser field missing", field=key)
    for key in allowed:
        if key in summary:
            projected[key] = summary[key]
    projected_packet = dict(packet)
    projected_packet["perf"] = {"corrected": projected}
    try:
        return base._compare_decoder_summary(projected_packet, parsed, dict(losses))
    except Exception as exc:  # noqa: BLE001
        _fail("decoder_summary_contract", f"{packet.get('packet_id')}: retained summary invariant failed: {exc}")


def _check_live_custody(packet_path: Path, packet: Mapping[str, Any], artifacts: Mapping[str, Mapping[str, Any]], plan_path: Path, plan: Mapping[str, Any]) -> dict[str, Any]:
    """Run retained 0977 custody and decoder-oracle checks for a live child."""
    _check_packet_plan_binding(packet, plan_path, plan)
    _check_packet_output_boundary(packet_path, packet, plan)
    base = _load_retained_0977_reducer()
    try:
        # _packet_artifacts validates every runner/decoder binding, including
        # command, terminal, report, loss, task-stream and final-log inputs.
        retained_artifacts = base._packet_artifacts(dict(packet), packet_path)
        launch, terminal = base._check_launch_terminal(dict(packet), retained_artifacts, dict(plan), packet_path)
        report = base._check_report(dict(packet), retained_artifacts, dict(plan))
        layout = base._check_layout(dict(packet), retained_artifacts, dict(plan))
        losses = base._check_loss(dict(packet), retained_artifacts)
        script_path = base._artifact_path(retained_artifacts, "corrected_script", "corrected.perf-script-no-inline.txt")
        parsed = base.p.parse_perf_script_strict(
            script_path.read_text(encoding="utf-8", errors="strict"),
            marker=plan["marker"]["symbol"], marker_dso=plan["binary"]["path"],
            expected_pid=int(launch["exec_pid"]), expected_cpu=int(plan["batch"]["cpu"]), expected_period=PERIOD,
        )
        summary = _compare_legacy_decoder_summary(base, packet, parsed, losses)
        log_path = base._artifact_path(retained_artifacts, "corrected_log", "corrected.dwfl.jsonl")
        log_stats = base._check_final_log(log_path, layout, plan["binary"]["path"])
    except Exception as exc:
        if isinstance(exc, ReducerFailure):
            raise
        _fail("custody_gate", f"{packet.get('packet_id')}: retained 0977 custody/oracle validation failed: {exc}")
    # The canonical 0978 namespace must point at the same files selected by
    # the retained checks.  This prevents a valid old artifact set from being
    # paired with a different analysis sidecar.
    for name in ("corrected_script", "corrected_log"):
        if name in artifacts and name in retained_artifacts:
            if check_binding(artifacts[name], f"{packet.get('packet_id')}.artifacts.{name}") != base._artifact_path(retained_artifacts, name):
                _fail("packet_binding", f"{packet.get('packet_id')}: canonical {name} is not the retained custody file")
    return {"artifacts": retained_artifacts, "launch": launch, "terminal": terminal, "report": report, "layout": layout, "loss": losses, "parsed": parsed, "summary": summary, "log_stats": log_stats}


def load_child(packet_path: Path, *, marker: str = DEFAULT_MARKER, marker_dso: str | None = None, allow_synthetic: bool = False, expected_timing_count: int | None = None, plan_path: Path | None = None, plan: Mapping[str, Any] | None = None) -> dict[str, Any]:
    packet_path = Path(packet_path).absolute()
    packet = load_json(packet_path, "packet")
    packet_id = packet.get("packet_id")
    if packet.get("schema") != PACKET_SCHEMA or packet.get("status") != "complete" or not isinstance(packet_id, str):
        _fail("packet_identity", "packet schema/status/identity is invalid", path=str(packet_path))
    child_number = packet.get("child_number")
    if not isinstance(child_number, int) or child_number != int(packet_id.rsplit("-", 1)[-1]):
        _fail("packet_identity", f"{packet_id}: child_number disagrees")
    artifacts = _artifact_map(packet)
    live_custody: dict[str, Any] | None = None
    if not allow_synthetic:
        if plan_path is None or not isinstance(plan, Mapping):
            _fail("plan_contract", f"{packet_id}: live child load lacks validated plan")
        _check_packet_output_boundary(packet_path, packet, plan)
        _check_child_artifact_boundaries(packet_path, artifacts, packet_id)
        live_custody = _check_live_custody(packet_path, packet, artifacts, Path(plan_path).absolute(), plan)
    digest_claims = packet.get("perf", {}).get("corrected", {}) if isinstance(packet.get("perf"), Mapping) and isinstance(packet.get("perf", {}).get("corrected"), Mapping) else {}
    for claim_name, artifact_name in (("corrected_script_sha256", "corrected_script"), ("corrected_log_sha256", "corrected_log")):
        if claim_name in digest_claims and digest_claims[claim_name] != artifacts[artifact_name].get("sha256"):
            _fail("packet_binding", f"{packet_id}: {claim_name} differs from bound artifact")
    corrected_path, corrected_obj = _read_artifact(artifacts, "corrected_script", f"{packet_id}.corrected_script")
    raw_path, raw_obj = _read_artifact(artifacts, "raw_stack", f"{packet_id}.raw_stack")
    lifecycle_path, lifecycle_obj = _read_artifact(artifacts, "lifecycle_ledger", f"{packet_id}.lifecycle_ledger")
    attribution_path, attribution_obj = _read_artifact(artifacts, "attribution_samples", f"{packet_id}.attribution_samples")
    if marker_dso is None:
        marker_dso = str(packet.get("marker_dso") or packet.get("binary", {}).get("path") or "")
    if not marker_dso:
        _fail("marker_contract", f"{packet_id}: marker DSO is missing")
    if corrected_obj is not None and not allow_synthetic:
        _fail("corrected_format", f"{packet_id}: live reducer requires the corrected perf-script text stream")
    if attribution_obj is not None and attribution_obj.get("schema") != "litchi.performance.0978.attribution_samples.v1":
        _fail("attribution_schema", f"{packet_id}: attribution_samples schema is not canonical", schema=attribution_obj.get("schema"))
    # The corrected perf script remains an independent identity source.  The
    # attribution sidecar supplies structured frames because pretty-printed
    # perf script is not a stable owner serialization.  The two streams must
    # still agree on every sample identity before the structured rows are used.
    corrected_script = _parse_corrected(corrected_path, corrected_obj, marker=marker, marker_dso=marker_dso)
    if attribution_obj is None:
        _fail("attribution_sidecar_format", f"{packet_id}: attribution_samples must be JSON")
    if attribution_obj.get("status") not in (None, "complete"):
        _fail("attribution_status", f"{packet_id}: attribution_samples is not complete", status=attribution_obj.get("status"))
    corrected_attribution_obj = {"sample_ledger": _rows_from_attribution(attribution_obj, "corrected_samples")}
    corrected = _parse_structured_rows(corrected_attribution_obj, "attribution.corrected_samples", marker=marker, marker_dso=marker_dso, allow_summary_claims=True)
    _compare_attribution_rows(corrected_script, corrected, "corrected_script", "attribution.corrected_samples")
    target_build_id = ""
    if isinstance(plan, Mapping) and isinstance(plan.get("marker"), Mapping):
        target_build_id = str(plan["marker"].get("build_id") or "").lower()
    _validate_exact_attribution_linkage(
        attribution_obj,
        corrected,
        marker=marker,
        marker_dso=marker_dso,
        target_build_id=target_build_id,
        allow_synthetic=allow_synthetic,
    )
    if raw_obj is None:
        _fail("raw_sidecar_format", f"{packet_id}: raw_stack must be JSON")
    try:
        raw = _parse_raw(raw_path, raw_obj, marker=marker, marker_dso=marker_dso)
    except ReducerFailure as error:
        if error.code != "sample_ledger_missing":
            raise
        raw = _parse_structured_rows({"sample_ledger": _rows_from_attribution(attribution_obj, "raw_samples")}, "attribution.raw_samples", marker=marker, marker_dso=marker_dso)
    attribution_raw = _parse_structured_rows({"sample_ledger": _rows_from_attribution(attribution_obj, "raw_samples")}, "attribution.raw_samples", marker=marker, marker_dso=marker_dso, allow_summary_claims=True)
    _compare_identities(raw, attribution_raw, "raw_stack", "attribution.raw_samples")
    scanned_raw = _reparse_raw_scanner(raw_path, raw_obj, artifacts, marker=marker, marker_dso=marker_dso, allow_synthetic=allow_synthetic)
    raw_scanner_rows: list[dict[str, Any]] | None = None
    if scanned_raw is not None:
        raw_scanner_rows, raw_scanner_original = scanned_raw
        _compare_identities(raw, raw_scanner_rows, "raw_stack", "independent raw scanner")
        raw = raw_scanner_rows
    else:
        raw_scanner_original = None
    lifecycle_obj_for_rows = dict(lifecycle_obj or {})
    if not isinstance(_first_mapping(lifecycle_obj_for_rows, ("sample_ledger", "samples", "rows")), list):
        if not allow_synthetic:
            _fail("lifecycle_contract", f"{packet_id}: lifecycle sidecar lacks its own full sample rows")
        lifecycle_obj_for_rows["sample_ledger"] = _rows_from_attribution(attribution_obj, "lifecycle_samples")
    lifecycle, ledger = _parse_lifecycle(lifecycle_path, lifecycle_obj_for_rows, marker=marker, marker_dso=marker_dso)
    if not allow_synthetic:
        _require_live_lifecycle_contract(lifecycle_obj, lifecycle, ledger)
    built_lifecycle = _reparse_lifecycle_stream(
        check_binding(artifacts["perf_mmap_dump"], f"{packet_id}.perf_mmap_dump"),
        raw_obj, lifecycle_obj, lifecycle,
        raw_scanner_rows=raw_scanner_original,
        attribution_obj=attribution_obj,
        packet=packet,
        plan=plan,
        marker=marker,
        marker_dso=marker_dso,
        allow_synthetic=allow_synthetic,
    )
    # The attribution sidecar is the canonical full-row handoff.  Reparse its
    # lifecycle stream independently of the lifecycle summary and compare it.
    attribution_lifecycle = _parse_structured_rows({"sample_ledger": _rows_from_attribution(attribution_obj, "lifecycle_samples")}, "attribution.lifecycle_samples", marker=marker, marker_dso=marker_dso, allow_summary_claims=True, allow_unframed_claims=True)
    _compare_lifecycle_rows(lifecycle, attribution_lifecycle, "lifecycle_ledger", "attribution.lifecycle_samples")
    lifecycle = attribution_lifecycle
    actual = _derive_child(corrected, raw, lifecycle, ledger, marker_dso=marker_dso)
    if built_lifecycle is not None:
        built_ledger = built_lifecycle.get("ledger")
        if not isinstance(built_ledger, Mapping):
            _fail("lifecycle_ledger_failure", f"{packet_id}: independent build_ledger returned no ledger")
        _compare_built_ledger_to_actual(built_ledger, actual)
    _compare_packet_claims(packet, _claim_object(packet), actual)
    gate = packet.get("gate")
    if not allow_synthetic:
        if not isinstance(packet.get("report"), Mapping) and "report" not in artifacts:
            _fail("report_missing", f"{packet_id}: report gate has no bound report")
        if "report" in artifacts and isinstance(packet.get("report"), Mapping):
            report_file = load_json(check_binding(artifacts["report"], f"{packet_id}.report"), f"{packet_id}.report")
            if report_file != packet["report"]:
                _fail("report_mismatch", f"{packet_id}: packet report differs from bound report artifact")
        if not isinstance(gate, Mapping) or gate.get("passed") is not True:
            _fail("packet_gate", f"{packet_id}: packet gate is not passing")
        if gate.get("raw_lost_sum") != 0 or gate.get("terminal_lost_samples") != 0:
            _fail("packet_gate", f"{packet_id}: packet loss gate is nonzero")
        if gate.get("passed_structural_conservation") is not True and gate.get("structural_conservation") is not True:
            _fail("packet_gate", f"{packet_id}: packet structural conservation gate is missing or failed")
        if gate.get("valid_empty_is_unknown") is not True:
            _fail("packet_gate", f"{packet_id}: packet does not retain valid empty samples as explicit unknown")
        if gate.get("marker_sample_count") is not None and gate.get("marker_sample_count") != actual["marker_sample_count"]:
            _fail("packet_gate", f"{packet_id}: packet marker count differs", packet=gate.get("marker_sample_count"), actual=actual["marker_sample_count"])
        minimum_marker = gate.get("minimum_marker_samples")
        if type(minimum_marker) is not int or actual["marker_sample_count"] < minimum_marker:
            _fail("packet_gate", f"{packet_id}: packet minimum marker gate is missing or failed", minimum=minimum_marker, actual=actual["marker_sample_count"])
    return {
        "packet_id": packet_id,
        "child_number": child_number,
        "packet_path": str(packet_path),
        "artifacts": artifacts,
        "actual": actual,
        "timing": _child_timing(packet, artifacts, expected_count=expected_timing_count),
        "marker": marker,
        "marker_dso": marker_dso,
    }


def _owner_bounds(actual: Mapping[str, Any]) -> dict[str, dict[str, float | int]]:
    t = int(actual["T"])
    unknown = int(actual["U_target"])
    owners = dict(actual["target_source_owner_period"])
    result: dict[str, dict[str, float | int]] = {}
    for owner in sorted(owners):
        known = int(owners[owner])
        result[owner] = {
            "K": known, "T": t, "U_target": unknown,
            "lower": known / t if t else 0.0,
            "upper": (known + unknown) / t if t else 0.0,
            "scope": "known physical-leaf period in target execution across all mapped DSOs; U_target is unresolved/unknown physical-owner period; bounds are per-owner and non-additive",
        }
    return result


def _correlations_for_owner(children: Sequence[Mapping[str, Any]], owner: str, key: str, t50: Sequence[float], t95: Sequence[float]) -> dict[str, Any]:
    values = [float(child["actual"].get(key, {}).get(owner, 0)) for child in children]
    result: dict[str, Any] = {"values_all_12": values, "vs_T50": _correlation(values, t50), "vs_T95": _correlation(values, t95)}
    if key == "target_known_owner_period":
        lower = [float(child["actual"].get("target_known_owner_period", {}).get(owner, 0)) / float(child["actual"]["T"]) if child["actual"]["T"] else 0.0 for child in children]
        upper = [(float(child["actual"].get("target_known_owner_period", {}).get(owner, 0)) + float(child["actual"]["U_target"])) / float(child["actual"]["T"]) if child["actual"]["T"] else 0.0 for child in children]
        result["lower_bounds_all_12"] = lower
        result["upper_bounds_all_12"] = upper
        result["lower_vs_T50"] = _correlation(lower, t50)
        result["lower_vs_T95"] = _correlation(lower, t95)
        result["upper_vs_T50"] = _correlation(upper, t50)
        result["upper_vs_T95"] = _correlation(upper, t95)
    return result


def reduce_children(children: Sequence[Mapping[str, Any]], *, plan_path: Path | None = None, manifest_path: Path | None = None) -> dict[str, Any]:
    if len(children) != CHILDREN:
        _fail("cohort_incomplete", f"reducer requires exactly {CHILDREN} children", actual=len(children))
    ordered = sorted(children, key=lambda row: int(row["child_number"]))
    expected = list(range(1, CHILDREN + 1))
    if [int(row["child_number"]) for row in ordered] != expected:
        _fail("cohort_identity", "child numbers are not exactly 1..12")
    for child in ordered:
        actual = child.get("actual")
        if not isinstance(actual, Mapping):
            _fail("period_conservation", "child has no derived period ledger", child=child.get("child_number"))
        if actual.get("D") != actual.get("M", 0) + actual.get("R", 0):
            _fail("period_conservation", "child D != M + R", child=child.get("child_number"), D=actual.get("D"), M=actual.get("M"), R=actual.get("R"))
        if actual.get("T") != actual.get("Tknown", 0) + actual.get("U_target", 0):
            _fail("period_conservation", "child T != Tknown + U_target", child=child.get("child_number"), T=actual.get("T"), Tknown=actual.get("Tknown"), U_target=actual.get("U_target"))
        if sum((actual.get("physical_leaf_owner_period") or {}).values()) + int(actual.get("physical_leaf_unknown_period", 0)) != int(actual.get("M", 0)):
            _fail("period_conservation", "child marker physical-leaf buckets do not conserve M", child=child.get("child_number"))
    t50 = [_percentile(child["timing"], 0.5) for child in ordered]
    t95 = [_percentile(child["timing"], 0.95) for child in ordered]
    metric_names = ("D", "M", "R", "T", "Tknown", "U_target", "U_marker")
    vectors = {name: [int(child["actual"][name]) for child in ordered] for name in metric_names}
    total_d = sum(vectors["D"])
    if total_d <= 0:
        _fail("period_conservation", "cohort D is not positive")
    children_output: list[dict[str, Any]] = []
    for child, p50, p95 in zip(ordered, t50, t95):
        actual = child["actual"]
        children_output.append({
            "packet_id": child["packet_id"], "child_number": child["child_number"],
            "timing": {"sample_count": len(child["timing"]), "p50_ns": p50, "p95_ns": p95, "p99_ns": _percentile(child["timing"], 0.99)},
            "period_metrics": {**{name: actual[name] for name in metric_names}, "R_label": "marker ancestry not observed", "D_share_of_all_children_period": actual["D"] / total_d},
            "sample_count": actual["sample_count"], "marker_sample_count": actual["marker_sample_count"],
            "physical_leaf_owner_period": actual["physical_leaf_owner_period"],
            "marker_known_owner_period": actual["marker_known_owner_period"],
            "physical_leaf_unknown_period": actual["physical_leaf_unknown_period"],
            "physical_leaf_unknown_samples": actual["physical_leaf_unknown_samples"],
            "inclusive_owner_period": actual["inclusive_owner_period"],
            "target_source_owner_period": actual["target_source_owner_period"],
            "target_known_owner_period": actual["target_known_owner_period"],
            "owner_bounds": _owner_bounds(actual),
            "phase_periods": actual["phase_periods"], "phase_counts": actual["phase_counts"],
            "artifacts": child["artifacts"],
        })
    owner_maps = {
        "marker_known_owner_period": sorted({owner for child in ordered for owner in child["actual"]["marker_known_owner_period"]}),
        "target_known_owner_period": sorted({owner for child in ordered for owner in child["actual"]["target_known_owner_period"]}),
        "target_source_owner_period": sorted({owner for child in ordered for owner in child["actual"]["target_source_owner_period"]}),
        "physical_leaf_owner_period": sorted({owner for child in ordered for owner in child["actual"]["physical_leaf_owner_period"]}),
        "inclusive_owner_period": sorted({owner for child in ordered for owner in child["actual"]["inclusive_owner_period"]}),
    }
    owner_associations: dict[str, Any] = {}
    for field, names in owner_maps.items():
        owner_associations[field] = {owner: _correlations_for_owner(ordered, owner, field, t50, t95) for owner in names}
    metric_correlations = {name: {"values_all_12": vectors[name], "vs_T50": _correlation([float(v) for v in vectors[name]], t50), "vs_T95": _correlation([float(v) for v in vectors[name]], t95)} for name in metric_names}
    variance_d = statistics.pvariance(vectors["D"])
    variance_m = statistics.pvariance(vectors["M"])
    variance_r = statistics.pvariance(vectors["R"])
    covariance_mr = sum((m - statistics.mean(vectors["M"])) * (r - statistics.mean(vectors["R"])) for m, r in zip(vectors["M"], vectors["R"])) / CHILDREN
    return {
        "schema": SCHEMA, "status": "complete", "analysis_only": True, "diagnostic_only": True,
        "qualification_claim": False, "speed_claim": False, "exact_timed_region_claim": False,
        "plan": {"path": str(plan_path.absolute()), "sha256": sha256(plan_path)} if plan_path else None,
        "manifest": {"path": str(manifest_path.absolute()), "sha256": sha256(manifest_path)} if manifest_path else None,
        "children": children_output,
        "summary": {
            "vectors": vectors, "timing_p50_ns": t50, "timing_p95_ns": t95,
            "metric_distributions": {name: _distribution(values) for name, values in vectors.items()},
            "metric_correlations": metric_correlations,
            "owner_associations": owner_associations,
            "period_conservation": {"D": sum(vectors["D"]), "M": sum(vectors["M"]), "R": sum(vectors["R"]), "D_equals_M_plus_R": sum(vectors["D"]) == sum(vectors["M"]) + sum(vectors["R"])},
            "variance_ledger": {"scope": "all 12 children; descriptive population moments", "variance_D": variance_d, "variance_M": variance_m, "variance_R": variance_r, "covariance_M_R": covariance_mr, "identity_passed": math.isclose(variance_d, variance_m + variance_r + 2 * covariance_mr, rel_tol=1e-12, abs_tol=1e-9)},
        },
        "attribution_contract": {
            "D": "all valid instructions:u period", "M": "period from marker-containing samples", "R": "marker ancestry not observed; may include missing/truncated ancestry, setup, or non-marker work",
            "T": "target execution period", "Tknown": "known physical-leaf period in target execution across all mapped DSOs", "U_target": "target execution period without a defensible physical owner",
            "U_marker": "marker period without a resolved physical leaf", "physical_leaf": "frame-0 owner across all mapped DSOs; disjoint with unknown",
            "inclusive": "deduplicated resolved marker-chain owners; overlapping and never summed", "owner_bounds": "per-owner [K/T,(K+U_target)/T]; bounds are non-additive",
            "scope": "all 12 children; descriptive diagnostic; no causal or speedup claim",
        },
    }


def load_cohort(manifest_path: Path, *, marker: str = DEFAULT_MARKER, marker_dso: str | None = None, plan_path: Path | None = None, allow_synthetic: bool = False) -> dict[str, Any]:
    manifest_path = Path(manifest_path).absolute()
    manifest = load_json(manifest_path, "manifest")
    if manifest.get("schema") != MANIFEST_SCHEMA or manifest.get("status") != "complete" or manifest.get("no_retry") is not True:
        _fail("cohort_incomplete", "manifest is not complete/no-retry")
    plan: dict[str, Any] | None = None
    if not allow_synthetic and plan_path is None:
        _fail("plan_contract", "live reduction requires --plan")
    paths = manifest.get("packets")
    if manifest.get("children_requested") != CHILDREN or manifest.get("children_completed") != CHILDREN or manifest.get("serial") is not True or not isinstance(paths, list) or len(paths) != CHILDREN or len(set(paths)) != CHILDREN:
        _fail("cohort_incomplete", "manifest does not contain exactly 12 unique packets")
    if plan_path is not None:
        plan_binding = manifest.get("plan")
        if not isinstance(plan_binding, Mapping) or plan_binding.get("path") != str(Path(plan_path).absolute()) or plan_binding.get("sha256") != sha256(plan_path):
            _fail("plan_binding", "manifest plan binding differs")
    if plan_path is not None:
        plan = _validate_plan(Path(plan_path).absolute())
        plan_marker = plan.get("marker") if isinstance(plan.get("marker"), Mapping) else {}
        if marker == DEFAULT_MARKER and isinstance(plan_marker.get("symbol"), str):
            marker = plan_marker["symbol"]
        if marker_dso is None:
            marker_dso = plan_marker.get("dso_path") or ((plan.get("binary") or {}).get("path") if isinstance(plan.get("binary"), Mapping) else None)
    children: list[dict[str, Any]] = []
    for expected, raw_path in enumerate(paths, 1):
        path = Path(raw_path).absolute()
        child = load_child(path, marker=marker, marker_dso=marker_dso, allow_synthetic=allow_synthetic, expected_timing_count=80 if plan is not None else None, plan_path=plan_path, plan=plan)
        if child["child_number"] is not None and child["child_number"] != expected:
            _fail("cohort_identity", "manifest order and packet child number differ", expected=expected, actual=child["child_number"])
        child["child_number"] = expected
        children.append(child)
    result = reduce_children(children, plan_path=plan_path, manifest_path=manifest_path)
    result["manifest_contract"] = {"children": CHILDREN, "serial": manifest.get("serial", True), "no_retry": True}
    return result


def _perf_ip_hex_regression() -> dict[str, Any]:
    """Guard both digits-only and alphabetic perf hexadecimal address tokens."""
    cases = {"591928": 0x591928, "e59f0": 0xE59F0}
    parsed = {}
    for token, expected in cases.items():
        actual = _parse_perf_ip_token(token, "self-test perf IP")
        if actual != expected:
            _fail("self_test", "perf hexadecimal IP regression failed", token=token, expected=expected, actual=actual)
        parsed[token] = actual
    return {"status": "pass", "cases": parsed, "base": 16}


def _synthetic_self_test() -> dict[str, Any]:
    # This exercises the real row normalizer, D/M/R and T conservation, and
    # the all-child reducer without touching a capture or accepting JSON as a
    # live packet input.
    hex_ip_regression = _perf_ip_hex_regression()
    target = {"dso": "/synthetic/target", "symbol_base": DEFAULT_MARKER, "resolved": True}
    base = [
        {"index": 1, "time_ns": 1, "period": PERIOD, "pid": 7, "tid": 7, "cpu": 31, "frames": [], "phase": "startup", "target_source_eligible": False},
        {"index": 2, "time_ns": 2, "period": PERIOD, "pid": 7, "tid": 7, "cpu": 31, "frames": [target], "phase": "target_execution", "target_source_eligible": True},
    ]
    rows = [_normalise_sample(row, index, "self-test", marker=DEFAULT_MARKER, marker_dso="/synthetic/target") for index, row in enumerate(base, 1)]
    actual = _derive_child(rows, rows, rows, {"D": 2 * PERIOD, "M": PERIOD, "R": PERIOD, "T": PERIOD, "U_target": 0, "U_marker": 0, "target_known_owner_period": {f"/synthetic/target::{DEFAULT_MARKER}": PERIOD}}, marker_dso="/synthetic/target")
    assert actual["D"] == 2 * PERIOD and actual["M"] == PERIOD and actual["R"] == PERIOD
    assert actual["T"] == actual["Tknown"] + actual["U_target"]
    children = [{"packet_id": f"child-{index:02d}", "child_number": index, "actual": actual, "timing": [float(index), float(index + 1)], "artifacts": {}} for index in range(1, CHILDREN + 1)]
    reduced = reduce_children(children)
    assert reduced["summary"]["vectors"]["D"] == [2 * PERIOD] * CHILDREN
    return {"schema": "litchi.performance.0978.reducer_self_test.v1", "status": "pass", "children": CHILDREN, "offline_only": True, "D": actual["D"], "M": actual["M"], "R": actual["R"], "perf_ip_hex_regression": hex_ip_regression}


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--self-test", action="store_true")
    parser.add_argument("--manifest", type=Path)
    parser.add_argument("--plan", type=Path)
    parser.add_argument("--output", type=Path)
    args = parser.parse_args(argv)
    try:
        if args.self_test:
            if args.manifest or args.plan or args.output:
                _fail("arguments", "--self-test is standalone")
            print(json.dumps(_synthetic_self_test(), indent=2, sort_keys=True))
            return 0
        if args.manifest is None or args.output is None or args.plan is None:
            _fail("arguments", "--plan, --manifest, and --output are required for live reduction")
        result = load_cohort(args.manifest, plan_path=args.plan)
        output = args.output.absolute()
        if output.exists():
            _fail("output_exists", "refusing to overwrite output", path=str(output))
        output.parent.mkdir(parents=True, exist_ok=True)
        output.write_text(json.dumps(result, indent=2, sort_keys=True) + "\n", encoding="utf-8")
        print(json.dumps({"status": "pass", "output": str(output), "sha256": sha256(output)}, sort_keys=True))
        return 0
    except ReducerFailure as exc:
        print(f"0978 reducer rejected: {exc}", file=sys.stderr)
        return 2
    except Exception as exc:
        print(f"0978 reducer rejected: internal {exc}", file=sys.stderr)
        return 2


__all__ = ["ReducerFailure", "SCHEMA", "PACKET_SCHEMA", "MANIFEST_SCHEMA", "CHILDREN", "file_binding", "sha256", "load_child", "load_cohort", "reduce_children"]


if __name__ == "__main__":
    raise SystemExit(main())
