#!/usr/bin/env python3
"""Decode one FP perf profile into conservative phase-ancestry buckets.

This consumes exports from the same perf.data. It never starts a workload and
never infers a timed phase from a leaf symbol alone. The candidate-review
mapping receipt owns all addresses, phase boundaries, binary identity, and
load-bias proof.
"""
from __future__ import annotations

import argparse
import collections
import hashlib
import json
import re
from pathlib import Path
from typing import Any

HEADER = re.compile(
    r"^\s*(?P<comm>\S+)\s+(?P<pid>\d+)/(?P<tid>\d+)\s+"
    r"(?P<time>[^:]+):\s+(?P<period>\d+)\s+(?P<event>[^:]+):\s*(?P<rest>.*)$"
)
HEX = re.compile(r"^(?:0x)?[0-9a-fA-F]+$")
DSO_PAREN = re.compile(r"^(?P<symbol>.*?)\s+\((?P<dso>(?:/|\[)[^)]*)\)$")
PHASES = ("query", "edit", "commit", "publication", "setup", "verification", "outside_profile", "unknown")
TIMED = {"query", "edit", "commit", "publication"}


def sha256(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for block in iter(lambda: f.read(1024 * 1024), b""):
            h.update(block)
    return h.hexdigest()


def as_int(value: Any) -> int | None:
    if isinstance(value, int):
        return value
    if isinstance(value, str) and HEX.fullmatch(value.strip()):
        return int(value, 16)
    return None


def parse_frame(text: str) -> dict[str, Any]:
    raw = text.strip()
    parts = raw.split(maxsplit=1)
    # A leaf-only perf script header carries the privilege suffix (`u:` for
    # user samples) before the IP.  Callchain lines do not carry that suffix.
    # Remove only the known perf privilege tokens; a symbol or DSO must never
    # be shifted heuristically.
    if parts and parts[0] in {"u:", "k:", "h:", "H:"}:
        raw = parts[1].strip() if len(parts) == 2 else ""
        parts = raw.split(maxsplit=1)
    token = parts[0] if parts else ""
    rest = parts[1] if len(parts) == 2 else ""
    dso = ""
    match = DSO_PAREN.match(rest)
    if match:
        symbol = match.group("symbol")
        dso = match.group("dso")
    else:
        symbol = rest
    return {
        "raw": raw,
        "ip": as_int(token),
        "symbol": symbol,
        "dso": dso,
        "unknown": "[unknown]" in raw.lower(),
    }


def event_key(row: dict[str, Any]) -> tuple[Any, ...]:
    return (row["comm"], row["pid"], row["tid"], row["time"], row["event"], row["period"])


def parse_export(path: Path, with_frames: bool) -> tuple[list[dict[str, Any]], dict[str, int]]:
    rows: list[dict[str, Any]] = []
    errors = collections.Counter()
    lines = path.read_text(encoding="utf-8", errors="replace").splitlines()
    current: dict[str, Any] | None = None
    for line in lines:
        match = HEADER.match(line)
        if match:
            if current is not None:
                rows.append(current)
            current = {
                "comm": match.group("comm"),
                "pid": int(match.group("pid")),
                "tid": int(match.group("tid")),
                "time": match.group("time"),
                "event": match.group("event").strip(),
                "period": int(match.group("period")),
                "leaf_frame": parse_frame(match.group("rest")),
                "frames": [],
            }
            continue
        if current is None:
            if line.strip():
                errors["orphan_nonempty_line"] += 1
            continue
        if with_frames and line.startswith("\t") and line.strip():
            current["frames"].append(parse_frame(line))
    if current is not None:
        rows.append(current)
    if with_frames:
        for row in rows:
            row["leaf_frame"] = row["frames"][0] if row["frames"] else parse_frame("")
    return rows, dict(errors)


def _append_call_boundary(boundaries: dict[str, dict[str, Any]], item: dict[str, Any]) -> None:
    """Convert one candidate-review callsite into the decoder's generic shape.

    The review map intentionally stores the source-level callsites rather than
    duplicating a decoder-specific phase-boundary format.  This adapter keeps
    the address set data-driven: no phase PC is embedded in this script.
    """
    phase_by_scope = {
        "query": "query",
        "edit": "edit",
        "commit": "commit",
        "publication": "publication",
        "setup": "setup",
        "verification": "verification",
        "expected_oracle": "outside_profile",
    }
    phase = phase_by_scope.get(item.get("scope"))
    if phase is None:
        raise ValueError(f"unrecognized phase-map scope: {item.get('scope')!r}")
    boundary = boundaries.setdefault(
        phase,
        {"phase": phase, "call_pcs": [], "return_pcs": [], "ranges": [], "callee_symbols": []},
    )
    for key in ("call_pc", "return_pc", "return_minus_one"):
        value = as_int(item.get(key))
        if value is None:
            raise ValueError(f"phase-map call lacks hexadecimal {key}: {item!r}")
        target = "call_pcs" if key == "call_pc" else "return_pcs"
        boundary[target].append(hex(value))
    # Timed phase claims need a workload-DSO callee in the same FP ancestry.
    # The raw review map's non-timer body call names are the authority; timer
    # calls and internal snapshot calls do not satisfy this gate by themselves.
    target = item.get("target")
    if (
        phase in TIMED
        and item.get("phase") == phase
        and isinstance(target, str)
        and not target.startswith("<")
    ):
        callee = target.split(" [", 1)[0]
        if callee not in boundary["callee_symbols"]:
            boundary["callee_symbols"].append(callee)


def adapt_candidate_review_calls(mapping: dict[str, Any]) -> dict[str, Any]:
    """Adapt the raw 0929 phase map without changing its on-disk identity."""
    if "phase_boundaries" in mapping:
        return mapping
    calls = mapping.get("calls")
    internal = mapping.get("publication_internal_snapshot", [])
    if not isinstance(calls, list) or not isinstance(internal, list):
        return mapping
    boundaries: dict[str, dict[str, Any]] = {}
    for item in calls + internal:
        if not isinstance(item, dict):
            raise ValueError("phase-map calls must contain JSON objects")
        _append_call_boundary(boundaries, item)
    # Keep a stable phase order in the derived in-memory representation.  The
    # original file remains the identity-bound input and is never rewritten.
    mapping["phase_boundaries"] = [
        boundaries[phase]
        for phase in PHASES
        if phase in boundaries
    ]
    return mapping


def load_mapping(path: Path) -> dict[str, Any]:
    mapping = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(mapping, dict):
        raise ValueError("mapping must be a JSON object")
    mapping = adapt_candidate_review_calls(mapping)
    required = ("schema", "binary", "phase_boundaries")
    missing = [key for key in required if key not in mapping]
    if missing:
        raise ValueError(f"mapping missing required fields: {', '.join(missing)}")
    if not isinstance(mapping["phase_boundaries"], list) or not mapping["phase_boundaries"]:
        raise ValueError("mapping phase_boundaries must be a nonempty list")
    if "load_bias" in mapping and as_int(mapping["load_bias"]) is None:
        raise ValueError("mapping load_bias is not a hexadecimal integer")
    return mapping


def mapping_identity(mapping: dict[str, Any], key: str) -> Any:
    direct = mapping.get(key)
    if direct is not None:
        return direct
    nested = mapping.get("workload")
    return nested.get(key) if isinstance(nested, dict) else None


def normalize(ip: int | None, dso: str, load_bias: int, workload_dso: str) -> int | None:
    # Never normalize an inherited or system DSO against the benchmark ELF.
    if ip is None or dso != workload_dso or ip < load_bias:
        return None
    return ip - load_bias


def values(boundary: dict[str, Any], key: str) -> set[int]:
    raw = boundary.get(key, [])
    if not isinstance(raw, list):
        return set()
    return {value for item in raw if (value := as_int(item)) is not None}


def ranges(boundary: dict[str, Any], key: str = "ranges") -> list[tuple[int, int]]:
    result: list[tuple[int, int]] = []
    raw = boundary.get(key, [])
    if not isinstance(raw, list):
        return result
    for item in raw:
        if isinstance(item, dict):
            start, end = as_int(item.get("start")), as_int(item.get("end"))
        elif isinstance(item, list) and len(item) == 2:
            start, end = as_int(item[0]), as_int(item[1])
        else:
            continue
        if start is not None and end is not None and start <= end:
            result.append((start, end))
    return result


def in_ranges(ip: int | None, rs: list[tuple[int, int]]) -> bool:
    return ip is not None and any(start <= ip <= end for start, end in rs)


def has_callee_ancestor(
    row: dict[str, Any], boundary: dict[str, Any], workload_dso: str
) -> bool:
    required = boundary.get("callee_symbols", [])
    if not required:
        return True
    return any(
        frame["dso"] == workload_dso
        and any(required_symbol in frame["symbol"] for required_symbol in required)
        for frame in row["frames"]
    )


def classify(
    row: dict[str, Any],
    mapping: dict[str, Any],
    load_bias: int,
    workload_comm: str,
    workload_dso: str,
) -> tuple[str, str]:
    if row["comm"] != workload_comm:
        return "outside_profile", "inherited_or_non_workload_comm"
    static_ips = [
        normalized
        for frame in row["frames"]
        if (normalized := normalize(frame["ip"], frame["dso"], load_bias, workload_dso)) is not None
    ]
    direct_ip = normalize(
        row["leaf_frame"]["ip"], row["leaf_frame"]["dso"], load_bias, workload_dso
    )
    matches: set[str] = set()
    direct_non_timed: set[str] = set()
    for boundary in mapping["phase_boundaries"]:
        phase = boundary.get("phase")
        if phase not in PHASES or phase == "unknown":
            continue
        return_pcs = values(boundary, "return_pcs")
        call_pcs = values(boundary, "call_pcs")
        if any(ip in return_pcs or ip in call_pcs for ip in static_ips):
            matches.add(phase)
        if phase in {"setup", "verification", "outside_profile"} and in_ranges(direct_ip, ranges(boundary)):
            direct_non_timed.add(phase)
    if len(matches) == 1:
        phase = next(iter(matches))
        timed_boundaries = [
            boundary
            for boundary in mapping["phase_boundaries"]
            if boundary.get("phase") == phase
        ]
        if phase in TIMED and timed_boundaries and not any(
            has_callee_ancestor(row, boundary, workload_dso)
            for boundary in timed_boundaries
        ):
            return "unknown", "mapped_timed_without_callee_ancestry"
        return phase, "mapped_frame_ancestry"
    if len(matches) > 1:
        return "unknown", "conflicting_phase_ancestry"
    if len(direct_non_timed) == 1:
        return next(iter(direct_non_timed)), "mapped_direct_non_timed_ip"
    if len(direct_non_timed) > 1:
        return "unknown", "conflicting_direct_non_timed_ip"
    if not row["frames"]:
        return "unknown", "empty_callchain"
    if any(frame["unknown"] or frame["ip"] is None for frame in row["frames"]):
        return "unknown", "invalid_or_unknown_ancestry"
    return "unknown", "unmapped_ancestry"


def symbol_key(frame: dict[str, Any]) -> str:
    symbol = frame["symbol"] or "[unknown]"
    # perf appends the sampled instruction offset to a resolved symbol.  The
    # phase report needs function-family totals, so retain the function name
    # and drop only a terminal `+0x...` suffix.  DSO identity remains part of
    # the key and unknown symbols are untouched.
    symbol = re.sub(r"\+0x[0-9a-fA-F]+$", "", symbol)
    dso = frame["dso"] or "[unknown-dso]"
    return f"{symbol} [{dso}]"


def top_symbols(rows: dict[str, list[int]], total_period: int) -> dict[str, Any]:
    ordered = sorted(rows.items(), key=lambda item: item[1][1], reverse=True)
    top = []
    remainder_samples = 0
    remainder_period = 0
    for index, (symbol, (samples, period)) in enumerate(ordered):
        if index < 15:
            top.append({
                "symbol": symbol,
                "samples": samples,
                "period": period,
                "within_phase_share": period / total_period if total_period else 0.0,
            })
        else:
            remainder_samples += samples
            remainder_period += period
    return {
        "top": top,
        "remainder": {
            "samples": remainder_samples,
            "period": remainder_period,
            "within_phase_share": remainder_period / total_period if total_period else 0.0,
        },
    }


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--leaf-export", type=Path, required=True)
    ap.add_argument("--stack-export", type=Path, required=True)
    ap.add_argument("--mapping", type=Path, required=True)
    ap.add_argument("--output", type=Path, required=True)
    ap.add_argument("--binary-sha256", required=True)
    ap.add_argument("--source-sha256", required=True)
    ap.add_argument(
        "--load-bias",
        help="capture-specific PIE load bias when the raw phase map deliberately omits it",
    )
    ap.add_argument(
        "--load-bias-receipt",
        type=Path,
        help="optional receipt that binds the supplied load bias to this perf.data",
    )
    ap.add_argument("--workload-comm")
    ap.add_argument("--workload-dso")
    args = ap.parse_args()

    mapping = load_mapping(args.mapping)
    errors: list[str] = []
    binary = mapping.get("binary", {})
    mapped_binary = binary.get("sha256") if isinstance(binary, dict) else None
    if mapped_binary != args.binary_sha256:
        errors.append("mapping binary SHA does not match requested profile binary SHA")
    source_field = mapping.get("source")
    source_sha_field = source_field.get("sha256") if isinstance(source_field, dict) else None
    mapped_source = mapping.get("source_sha256") or source_sha_field
    if mapped_source != args.source_sha256:
        errors.append("mapping source SHA does not match requested source SHA")
    load_bias = as_int(mapping.get("load_bias"))
    cli_load_bias = as_int(args.load_bias)
    if load_bias is None:
        load_bias = cli_load_bias
    elif cli_load_bias is not None and cli_load_bias != load_bias:
        errors.append("CLI load bias does not match mapping load bias")
    if load_bias is None:
        errors.append("capture-specific load bias is required")
        load_bias = 0
    load_bias_receipt: dict[str, Any] | None = None
    if args.load_bias_receipt is not None:
        try:
            load_bias_receipt = json.loads(args.load_bias_receipt.read_text(encoding="utf-8"))
            if not isinstance(load_bias_receipt, dict):
                errors.append("load-bias receipt is not a JSON object")
                load_bias_receipt = None
            elif load_bias_receipt.get("load_bias") != hex(load_bias):
                errors.append("load-bias receipt does not match decoder load bias")
        except (OSError, json.JSONDecodeError) as exc:
            errors.append(f"load-bias receipt could not be read: {exc}")
    workload_comm = args.workload_comm or mapping_identity(mapping, "comm")
    workload_dso = args.workload_dso or mapping_identity(mapping, "dso")
    if not isinstance(workload_comm, str) or not workload_comm:
        errors.append("workload comm is required to distinguish inherited events")
        workload_comm = ""
    if not isinstance(workload_dso, str) or not workload_dso:
        errors.append("workload DSO is required to normalize frame PCs")
        workload_dso = ""

    leaf, leaf_parse_errors = parse_export(args.leaf_export, with_frames=False)
    stack, stack_parse_errors = parse_export(args.stack_export, with_frames=True)
    leaf_queues: dict[tuple[Any, ...], collections.deque[dict[str, Any]]] = collections.defaultdict(collections.deque)
    stack_queues: dict[tuple[Any, ...], collections.deque[dict[str, Any]]] = collections.defaultdict(collections.deque)
    for row in leaf:
        leaf_queues[event_key(row)].append(row)
    for row in stack:
        stack_queues[event_key(row)].append(row)
    duplicate_keys = sum(len(q) - 1 for q in leaf_queues.values() if len(q) > 1)
    duplicate_keys += sum(len(q) - 1 for q in stack_queues.values() if len(q) > 1)
    joined: list[dict[str, Any]] = []
    missing_stack = 0
    for key, queue in leaf_queues.items():
        while queue:
            leaf_row = queue.popleft()
            stack_queue = stack_queues.get(key)
            if not stack_queue:
                missing_stack += 1
                continue
            stack_row = stack_queue.popleft()
            leaf_row["frames"] = stack_row["frames"]
            leaf_row["stack_leaf_frame"] = stack_row["leaf_frame"]
            joined.append(leaf_row)
    missing_leaf = sum(len(q) for q in stack_queues.values())

    ancestry = collections.Counter()
    reasons = collections.Counter()
    buckets = {
        phase: {"samples": 0, "period": 0, "symbols": collections.defaultdict(lambda: [0, 0])}
        for phase in PHASES
    }
    for row in joined:
        if row["comm"] != workload_comm:
            ancestry["inherited_events"] += 1
        frames = row["frames"]
        if frames:
            ancestry["nonempty_callchains"] += 1
            stack_leaf = row["stack_leaf_frame"]
            leaf_frame = row["leaf_frame"]
            if leaf_frame["ip"] is not None and stack_leaf["ip"] is not None:
                ancestry["leaf_first_frame_comparable"] += 1
                if leaf_frame["ip"] == stack_leaf["ip"] and leaf_frame["dso"] == stack_leaf["dso"]:
                    ancestry["leaf_first_frame_ip_dso_match"] += 1
                else:
                    ancestry["leaf_first_frame_ip_dso_mismatch"] += 1
            if any(frame["unknown"] for frame in frames):
                ancestry["unknown_frames"] += sum(frame["unknown"] for frame in frames)
            ancestry["invalid_frame_addresses"] += sum(frame["ip"] is None for frame in frames)
            ancestry["non_workload_dso_frames"] += sum(
                bool(frame["dso"]) and frame["dso"] != workload_dso for frame in frames
            )
            ancestry["normalized_workload_frames"] += sum(
                normalize(frame["ip"], frame["dso"], load_bias, workload_dso) is not None for frame in frames
            )
        else:
            ancestry["empty_callchains"] += 1
        phase, reason = classify(row, mapping, load_bias, workload_comm, workload_dso)
        buckets[phase]["samples"] += 1
        buckets[phase]["period"] += row["period"]
        symbol = symbol_key(row["leaf_frame"])
        buckets[phase]["symbols"][symbol][0] += 1
        buckets[phase]["symbols"][symbol][1] += row["period"]
        reasons[reason] += 1
        if phase in TIMED:
            ancestry[f"proven_{phase}_samples"] += 1
            ancestry[f"proven_{phase}_period"] += row["period"]

    leaf_period = sum(row["period"] for row in leaf)
    stack_period = sum(row["period"] for row in stack)
    joined_period = sum(row["period"] for row in joined)
    if missing_stack or missing_leaf or len(leaf) != len(stack) or leaf_period != stack_period or joined_period != leaf_period:
        errors.append("leaf/stack join or period conservation failed")

    result_buckets: dict[str, Any] = {}
    for phase, data in buckets.items():
        symbols = {key: value for key, value in data["symbols"].items()}
        result_buckets[phase] = {
            "samples": data["samples"],
            "period": data["period"],
            "full_profile_period_share": data["period"] / joined_period if joined_period else 0.0,
            "leaf_symbols": top_symbols(symbols, data["period"]),
        }
    result = {
        "schema": "litchi.performance.0929.fp-phase-decoder.v2",
        "status": "failed_closed" if errors else "descriptive_pass",
        "claim_scope": "whole-process FP ancestry description only; no timing or speedup claim",
        "inputs": {
            "leaf_export": {"path": str(args.leaf_export), "bytes": args.leaf_export.stat().st_size, "sha256": sha256(args.leaf_export)},
            "stack_export": {"path": str(args.stack_export), "bytes": args.stack_export.stat().st_size, "sha256": sha256(args.stack_export)},
            "mapping": {"path": str(args.mapping), "bytes": args.mapping.stat().st_size, "sha256": sha256(args.mapping)},
            "binary_sha256": args.binary_sha256,
            "source_sha256": args.source_sha256,
            "workload_comm": workload_comm,
            "workload_dso": workload_dso,
            "load_bias": hex(load_bias),
            "load_bias_receipt": (
                {
                    "path": str(args.load_bias_receipt),
                    "bytes": args.load_bias_receipt.stat().st_size,
                    "sha256": sha256(args.load_bias_receipt),
                }
                if args.load_bias_receipt is not None and args.load_bias_receipt.exists()
                else None
            ),
            "phase_map_adapter": "raw calls plus publication_internal_snapshot grouped by reviewed scope in memory",
        },
        "conservation": {
            "leaf_samples": len(leaf),
            "stack_samples": len(stack),
            "joined_samples": len(joined),
            "missing_stack": missing_stack,
            "missing_leaf": missing_leaf,
            "duplicate_keys": duplicate_keys,
            "leaf_period": leaf_period,
            "stack_period": stack_period,
            "joined_period": joined_period,
            "leaf_parse_errors": leaf_parse_errors,
            "stack_parse_errors": stack_parse_errors,
        },
        "ancestry_quality": dict(ancestry),
        "classification_reasons": dict(reasons),
        "buckets": result_buckets,
        "errors": errors,
    }
    args.output.write_text(json.dumps(result, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    print(json.dumps({"status": result["status"], "errors": errors, "output": str(args.output)}))
    return 0 if not errors else 2


if __name__ == "__main__":
    raise SystemExit(main())
