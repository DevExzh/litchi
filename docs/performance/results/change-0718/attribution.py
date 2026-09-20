#!/usr/bin/env python3
"""Small, read-only parsers for the 0718 attribution tools.

The capture driver keeps the tool output beside its receipt.  This module only
decodes one retained ``strace`` or DHAT file; it deliberately does not run a
capture, inspect the repository, or write an audit result.  The returned
objects contain ordinary JSON values so the caller can put them in a retained
analysis document without a second parser.

There are two deliberately separate notions of ownership here.  A stack is a
``phase`` stack only when both owner markers are visible.  A stack containing
``write_plain`` but not ``ordinary_save::run_case`` is reported as
``setup_publication``.  Missing, unknown, or frame-limit-truncated context is
always retained as ``other_unresolved`` and never proves that an owner was
absent.
"""

from __future__ import annotations

import ast
import json
import re
from pathlib import Path
from typing import Any, Iterable, Mapping, Sequence


SCHEMA_VERSION = 1
GROUPS = ("phase", "setup_publication", "other_unresolved")
DEFAULT_SYSCALLS = ("mmap", "mremap", "munmap", "brk")
DEFAULT_FRAME_LIMIT = 40
DEFAULT_DHAT_FRAME_LIMIT = 40
CONTROL_PAIRS = 32
PROC_FILES = ("io", "stat", "status")
MAPPING_SYSCALLS = frozenset(("mmap", "mremap", "munmap", "brk", "madvise", "mprotect"))


class AttributionError(ValueError):
    """A missing, malformed, or contradictory profiler artifact."""


def _require(condition: bool, message: str) -> None:
    if not condition:
        raise AttributionError(message)


def _read_text(path: Path) -> str:
    path = Path(path)
    _require(path.is_file() and not path.is_symlink(), f"missing regular file: {path}")
    try:
        return path.read_text(encoding="utf-8", errors="replace")
    except OSError as error:
        raise AttributionError(f"cannot read {path}: {error}") from error


def _read_json(path: Path) -> Any:
    try:
        return json.loads(_read_text(Path(path)))
    except json.JSONDecodeError as error:
        raise AttributionError(f"invalid JSON in {path}: {error}") from error


def _mapping_settings(plan: Mapping[str, Any]) -> Mapping[str, Any]:
    lanes = plan.get("lanes", {})
    settings = lanes.get("mapping", {}) if isinstance(lanes, Mapping) else {}
    _require(isinstance(settings, Mapping), "plan mapping lane is malformed")
    return settings


def _allocation_settings(plan: Mapping[str, Any]) -> Mapping[str, Any]:
    lanes = plan.get("lanes", {})
    settings = lanes.get("allocation", {}) if isinstance(lanes, Mapping) else {}
    _require(isinstance(settings, Mapping), "plan allocation lane is malformed")
    return settings


def _options(settings: Mapping[str, Any]) -> list[str]:
    options = settings.get("options", [])
    _require(isinstance(options, list) and all(isinstance(v, str) for v in options),
             "tool options in plan are malformed")
    return list(options)


def _trace_syscalls(plan: Mapping[str, Any]) -> tuple[str, ...]:
    settings = _mapping_settings(plan)
    names: list[str] = []
    for option in _options(settings):
        if option.startswith("trace="):
            names.extend(part for part in option[6:].split(",") if part)
    if not names:
        names.extend(DEFAULT_SYSCALLS)
    # Keep plan order while rejecting duplicate spellings.
    result = tuple(dict.fromkeys(names))
    _require(all(re.fullmatch(r"[A-Za-z_][A-Za-z0-9_]*", name) for name in result),
             "plan trace syscall name is malformed")
    return result


def _frame_limit(options: Sequence[str], default: int) -> int:
    # strace uses the singular spelling.  Accept the old spelling as a parser
    # compatibility measure because it is harmless and makes malformed pilot
    # artifacts fail with a useful limit rather than an unrelated error.
    for option in options:
        match = re.fullmatch(r"--stack-trace-frame-limit=(\d+)", option)
        if match is None:
            match = re.fullmatch(r"--stack-traces-frame-limit=(\d+)", option)
        if match is not None:
            value = int(match.group(1))
            _require(value > 0, "stack trace frame limit must be positive")
            return value
    return default


def _owner_frames(plan: Mapping[str, Any]) -> tuple[str, str]:
    value = plan.get("owner_frames", ["write_plain", "ordinary_save::run_case"])
    _require(isinstance(value, list) and all(isinstance(v, str) and v for v in value),
             "plan owner_frames is malformed")
    _require(len(value) >= 2, "plan must name write_plain and ordinary_save::run_case")
    write = next((v for v in value if "write_plain" in v), None)
    run_case = next((v for v in value if "ordinary_save::run_case" in v), None)
    _require(write is not None and run_case is not None,
             "plan owner_frames must include write_plain and ordinary_save::run_case")
    return write, run_case


def _contains_marker(stack: Sequence[str], marker: str) -> bool:
    return marker in "\n".join(stack)


def _unknown_frame(frame: str) -> bool:
    lower = frame.strip().lower()
    if not lower:
        return True
    # strace prints a stripped executable as ``/path/procfs()``.  The address
    # is present, but the symbol is not, so treating this as resolved would
    # turn a missing owner name into a false owner-absence claim.
    if re.search(r"\(\s*\)", lower):
        return True
    return any(token in lower for token in ("???", "[unknown]", "<unknown>", "unknown"))


def _truncated_frame(frame: str) -> bool:
    text = frame.strip().lower()
    return text in {"...", "<truncated>", "[truncated]"} or "truncated" in text


def _context_status(stack: Sequence[str], limit: int) -> dict[str, Any]:
    unknown_frames = sum(1 for frame in stack if _unknown_frame(frame))
    marker_truncation = any(_truncated_frame(frame) for frame in stack)
    # A stack ending at the configured limit is evidence of the frame-limit
    # boundary, not evidence that the call chain ended there.
    truncated = marker_truncation or (limit > 0 and len(stack) >= limit)
    unknown_context = not stack or unknown_frames > 0
    return {
        "frame_count": len(stack),
        "unknown_frame_count": unknown_frames,
        "unknown_context": unknown_context,
        "truncated_context": truncated,
    }


ADDRESS_RE = re.compile(r"\[(0x[0-9A-Fa-f]+)\]")


def _frame_parts(frame: str) -> tuple[str | None, str | None, str | None]:
    """Return (object path, in-frame symbol, raw address) for a strace frame."""

    address_match = ADDRESS_RE.search(frame)
    address = address_match.group(1) if address_match else None
    prefix, separator, _ = frame.partition("(")
    if not separator:
        return None, None, address
    object_path = prefix.strip()
    symbol = frame[len(prefix) + 1:].split(")", 1)[0].strip()
    return object_path or None, symbol or None, address


def _symbolization_binding(symbolization: Mapping[str, Any]) -> tuple[str, Mapping[str, Any]]:
    frames = symbolization.get("frames")
    _require(isinstance(frames, Mapping), "symbolization frames are malformed")
    binary = symbolization.get("binary")
    binary_path: Any = symbolization.get("binary_path")
    if binary_path is None and isinstance(binary, Mapping):
        binary_path = binary.get("path")
    if binary_path is None and isinstance(binary, str):
        binary_path = binary
    if binary_path is None:
        binary_path = symbolization.get("path")
    _require(isinstance(binary_path, str) and binary_path,
             "symbolization binary path is missing")
    return binary_path, frames


def _symbolization_frame(frames: Mapping[str, Any], address: str) -> Mapping[str, Any] | None:
    # The retained witness uses the exact lower-case hex spelling from strace.
    # A couple of case variants are accepted for hand-built parser fixtures,
    # while the address value itself is never rounded or range-matched.
    candidates = (address, address.lower(), address.upper(),
                  "0x" + address[2:].lower(), "0x" + address[2:].upper())
    for key in candidates:
        value = frames.get(key)
        if isinstance(value, Mapping):
            return value
    return None


def _resolve_mapping_stacks(events: Sequence[dict[str, Any]],
                            symbolization: Mapping[str, Any] | None,
                            frame_limit: int) -> dict[str, Any]:
    """Attach raw/resolved frame evidence and return its binding summary."""

    if symbolization is None:
        for event in events:
            event["classification_stack"] = list(event["stack"])
            event["stack_evidence"] = [
                {"raw": frame, "address": _frame_parts(frame)[2], "resolved": None}
                for frame in event["stack"]
            ]
        return {"provided": False, "status": "raw_only"}

    _require(isinstance(symbolization, Mapping), "symbolization is not an object")
    binary_path, frames = _symbolization_binding(symbolization)
    mapped = 0
    unresolved_local = 0
    seen_local: set[str] = set()
    for event in events:
        evidence: list[dict[str, Any]] = []
        classification_stack: list[str] = []
        for raw in event["stack"]:
            object_path, raw_symbol, address = _frame_parts(raw)
            resolved: str | None = None
            resolution = "raw_symbol" if raw_symbol and raw_symbol not in {"???"} else "unresolved"
            if address is not None and object_path == binary_path:
                seen_local.add(address.lower())
                entry = _symbolization_frame(frames, address)
                _require(entry is not None,
                         f"symbolization map has no exact frame entry for {address} in {binary_path}")
                entry_object = entry.get("binary_path", entry.get("object"))
                if entry_object is not None:
                    _require(entry_object == binary_path,
                             f"symbolization frame {address} binary path differs")
                function = entry.get("function")
                _require(isinstance(function, str) and function,
                         f"symbolization frame {address} has no function")
                resolved = function
                resolution = "symbolization_map"
                mapped += 1
            elif address is None and object_path == binary_path:
                # This branch is kept explicit so a future frame parser cannot
                # silently skip a local frame with an unparsable address.
                unresolved_local += 1
            elif raw_symbol and raw_symbol not in {"???"}:
                resolved = raw_symbol
            if resolved is not None:
                classification_stack.append(resolved)
            else:
                classification_stack.append(raw)
            evidence.append({
                "raw": raw,
                "resolved": resolved,
            })
        event["classification_stack"] = classification_stack
        event["stack_evidence"] = evidence
        unknown_frames = sum(
            1 for item in evidence
            if item["resolved"] is None and _unknown_frame(str(item["raw"]))
        )
        event["context_status"] = {
            "frame_count": len(evidence),
            "unknown_frame_count": unknown_frames,
            "unknown_context": not evidence or unknown_frames > 0,
            "truncated_context": (
                len(evidence) >= frame_limit
                or any(_truncated_frame(str(item["raw"])) for item in evidence)
            ),
        }
        event["unknown_context"] = event["context_status"]["unknown_context"]
        event["truncated_context"] = event["context_status"]["truncated_context"]
    return {
        "provided": True,
        "status": "mapped",
        "binary_path": binary_path,
        "frame_entry_count": len(frames),
        "mapped_frame_count": mapped,
        "unresolved_local_frame_count": unresolved_local,
        "unique_local_frame_count": len(seen_local),
    }


def _owner_evidence(stack: Sequence[str], status: Mapping[str, Any],
                    owner: tuple[str, str]) -> dict[str, str]:
    result: dict[str, str] = {}
    for label, marker in (("write_plain", owner[0]), ("ordinary_save::run_case", owner[1])):
        if _contains_marker(stack, marker):
            result[label] = "present"
        elif bool(status["unknown_context"]) or bool(status["truncated_context"]):
            result[label] = "unknown"
        else:
            result[label] = "absent_in_resolved_frames"
    return result


def _classify(stack: Sequence[str], owner: tuple[str, str]) -> str:
    has_write = _contains_marker(stack, owner[0])
    has_run_case = _contains_marker(stack, owner[1])
    if has_write and has_run_case:
        return "phase"
    if has_write and not has_run_case:
        return "setup_publication"
    return "other_unresolved"


def _split_args(value: str) -> list[str]:
    """Split a strace argument list without splitting nested/quoted commas."""

    result: list[str] = []
    start = 0
    depth = 0
    quote: str | None = None
    escaped = False
    for index, char in enumerate(value):
        if quote is not None:
            if escaped:
                escaped = False
            elif char == "\\":
                escaped = True
            elif char == quote:
                quote = None
            continue
        if char in "'\"":
            quote = char
        elif char in "([{":
            depth += 1
        elif char in ")]}":
            depth = max(depth - 1, 0)
        elif char == "," and depth == 0:
            result.append(value[start:index].strip())
            start = index + 1
    result.append(value[start:].strip())
    return result if result != [""] else []


def _integer(value: str) -> int | None:
    token = value.strip()
    if token in {"NULL", "null", "-", "?"}:
        return None
    match = re.fullmatch(r"[-+]?(?:0[xX][0-9a-fA-F]+|[0-9]+)", token)
    if match is None:
        return None
    try:
        base = 16 if token.lower().lstrip("+-").startswith("0x") else 10
        return int(token, base)
    except ValueError:
        return None


def _request_length(name: str, args: Sequence[str]) -> int | None:
    # For mremap, the requested extent is the new size.  For the other
    # extent-bearing calls, it is the second argument.  brk/openat have no
    # byte-length request and intentionally retain null.
    position = {"mmap": 1, "mremap": 2, "munmap": 1,
                "madvise": 1, "mprotect": 1}.get(name)
    if position is None or position >= len(args):
        return None
    value = _integer(args[position])
    return value if value is not None and value >= 0 else None


def _quoted_path(value: str) -> str | None:
    match = re.search(r'("(?:\\.|[^"\\])*")', value)
    if match is None:
        return None
    try:
        decoded = ast.literal_eval(match.group(1))
    except (SyntaxError, ValueError):
        return match.group(1)[1:-1]
    return decoded if isinstance(decoded, str) else None


def _proc_file(path: str | None) -> str | None:
    if path is None:
        return None
    match = re.search(r"/proc/(?:self|[0-9]+)/(?P<file>io|stat|status)(?:$|/)", path)
    return match.group("file") if match else None


def _terminal(line: str) -> int | None:
    match = re.fullmatch(r"\s*\+\+\+\s+exited\s+with\s+(\d+)\s+\+\+\+\s*", line)
    return int(match.group(1)) if match else None


def _is_killed_terminal(line: str) -> bool:
    return bool(re.fullmatch(r"\s*\+\+\+\s+killed\s+by\s+.+\+\+\+\s*", line))


SYSCALL_RE = re.compile(
    r"^\s*(?:(?P<pid>[0-9]+)\s+)?(?P<name>[A-Za-z_][A-Za-z0-9_]*)"
    r"\((?P<args>.*)\)\s+=\s+(?P<result>.+?)\s*$"
)
FRAME_RE = re.compile(r"^\s*>\s?(?P<frame>.*)$")


def _parse_syscall(line: str, allowed: set[str]) -> dict[str, Any] | None:
    match = SYSCALL_RE.match(line)
    if match is None or match.group("name") not in allowed:
        return None
    name = match.group("name")
    args_text = match.group("args")
    args = _split_args(args_text)
    request_path = _quoted_path(args_text) if name in {"open", "openat", "openat2"} else None
    return {
        "syscall": name,
        "pid": int(match.group("pid")) if match.group("pid") else None,
        "arguments": args_text,
        "result": match.group("result").strip(),
        "request_length": _request_length(name, args),
        "request_path": request_path,
        "stack": [],
    }


def _success_result(result: str) -> bool:
    return not result.lstrip().startswith("-1")


def _totals(events: Iterable[Mapping[str, Any]], syscall_names: Sequence[str]) -> dict[str, Any]:
    by_syscall: dict[str, dict[str, Any]] = {}
    for name in syscall_names:
        by_syscall[name] = {
            "request_count": 0,
            "successful_count": 0,
            "failed_count": 0,
            "known_length_count": 0,
            "request_length_sum": 0,
        }
    event_count = 0
    successful_count = 0
    failed_count = 0
    known_length_count = 0
    length_sum = 0
    for event in events:
        name = str(event["syscall"])
        row = by_syscall.setdefault(name, {
            "request_count": 0,
            "successful_count": 0,
            "failed_count": 0,
            "known_length_count": 0,
            "request_length_sum": 0,
        })
        event_count += 1
        row["request_count"] += 1
        if _success_result(str(event["result"])):
            successful_count += 1
            row["successful_count"] += 1
        else:
            failed_count += 1
            row["failed_count"] += 1
        length = event.get("request_length")
        if length is not None:
            known_length_count += 1
            length_sum += int(length)
            row["known_length_count"] += 1
            row["request_length_sum"] += int(length)
    return {
        "request_count": event_count,
        "successful_count": successful_count,
        "failed_count": failed_count,
        "known_length_count": known_length_count,
        "request_length_sum": length_sum,
        # This alias makes the unit explicit for callers that aggregate
        # mmap-like requests with the allocation parser.
        "requested_bytes": length_sum,
        "by_syscall": by_syscall,
    }


def _mapping_group(group: str, events: Sequence[Mapping[str, Any]],
                   syscall_names: Sequence[str], frame_limit: int,
                   owner: tuple[str, str]) -> dict[str, Any]:
    selected = [event for event in events if event["classification"] == group]
    details: list[dict[str, Any]] = []
    unknown_context_count = 0
    truncated_context_count = 0
    for event in selected:
        stack = list(event["stack"])
        classification_stack = list(event.get("classification_stack", stack))
        status = event.get("context_status")
        if not isinstance(status, Mapping):
            status = _context_status(stack, frame_limit)
        if status["unknown_context"]:
            unknown_context_count += 1
        if status["truncated_context"]:
            truncated_context_count += 1
        detail: dict[str, Any] = {
            "event_index": event["index"],
            "syscall": event["syscall"],
            "request_length": event["request_length"],
            "request_path": event["request_path"],
            "result": event["result"],
            "owner_evidence": _owner_evidence(classification_stack, status, owner),
            "unknown_context": status["unknown_context"],
            "truncated_context": status["truncated_context"],
            "stack_evidence": event.get("stack_evidence", []),
        }
        details.append(detail)
    totals = _totals(selected, syscall_names)
    return {
        "classification": group,
        "event_indices": [event["index"] for event in selected],
        "events": details,
        "totals": totals,
        "unknown_context_count": unknown_context_count,
        "truncated_context_count": truncated_context_count,
        "owner_absence_inference": "forbidden_when_context_is_missing_unknown_or_truncated",
    }


def _report_value(report: Any) -> Any:
    if isinstance(report, (str, Path)):
        return _read_json(Path(report))
    return report


def _report_result(report: Mapping[str, Any]) -> Mapping[str, Any]:
    results = report.get("results")
    if isinstance(results, list):
        _require(len(results) == 1, "raw report must contain exactly one result")
        _require(isinstance(results[0], Mapping), "raw report result is malformed")
        return results[0]
    return report


def _raw_sample_deltas(report: Any) -> tuple[list[Any], Mapping[str, Any]]:
    value = _report_value(report)
    _require(isinstance(value, Mapping), "raw report is not an object")
    result = _report_result(value)
    source = result.get("source")
    _require(isinstance(source, Mapping), "raw report source is missing")
    ordinary = source.get("ordinary_save")
    _require(isinstance(ordinary, Mapping), "raw report ordinary_save source is missing")
    probe = ordinary.get("process_probe")
    _require(isinstance(probe, Mapping), "raw report process_probe is missing")
    samples = probe.get("sample_deltas")
    _require(isinstance(samples, list), "raw report process_probe sample_deltas is malformed")
    return list(samples), result


def _marker_windows(events: Sequence[Mapping[str, Any]], plan: Mapping[str, Any],
                    report: Any, frame_limit: int,
                    owner: tuple[str, str]) -> dict[str, Any]:
    settings = _mapping_settings(plan)
    warmup = settings.get("warmup")
    samples = settings.get("samples")
    _require(isinstance(warmup, int) and warmup >= 0, "mapping warmup count is malformed")
    _require(isinstance(samples, int) and samples >= 0, "mapping sample count is malformed")

    candidates: list[Mapping[str, Any]] = []
    unresolved_candidates = 0
    for event in events:
        if event["syscall"] != "openat":
            continue
        proc_file = _proc_file(event.get("request_path"))
        if proc_file is None:
            continue
        candidates.append(event)
        if not _contains_marker(event.get("classification_stack", event["stack"]), owner[1]):
            unresolved_candidates += 1

    expected_snapshots = 2 * (CONTROL_PAIRS + warmup + samples)
    expected_marker_events = expected_snapshots * len(PROC_FILES)

    # A process can read /proc/self/status for a final cleanup or scheduler
    # diagnostic outside run_case.  It is retained as an extra event, while
    # only complete io/stat/status triplets participate in the protocol.  A
    # malformed event inside a purported owner triplet remains a hard error;
    # this keeps the window alignment fail-closed.
    triplets: list[list[Mapping[str, Any]]] = []
    extras: list[Mapping[str, Any]] = []
    cursor = 0
    while cursor < len(candidates):
        paths = [_proc_file(candidates[i].get("request_path"))
                 for i in range(cursor, min(cursor + len(PROC_FILES), len(candidates)))]
        if len(paths) == len(PROC_FILES) and paths == list(PROC_FILES):
            triplets.append(candidates[cursor:cursor + len(PROC_FILES)])
            cursor += len(PROC_FILES)
            continue
        current = candidates[cursor]
        if (_proc_file(current.get("request_path")) == "status"
                and not _contains_marker(
                    current.get("classification_stack", current["stack"]), owner[1]
                )):
            extras.append(current)
            cursor += 1
            continue
        raise AttributionError(
            f"malformed procfs marker triplet near event {current['index']}"
        )
    _require(len(triplets) == expected_snapshots,
             f"expected {expected_snapshots} procfs snapshots, got {len(triplets)}")

    marker_events = [event for triplet in triplets for event in triplet]
    unresolved_candidates = sum(
        not _contains_marker(event.get("classification_stack", event["stack"]), owner[1])
        for event in marker_events
    )

    snapshots: list[dict[str, Any]] = []
    for ordinal, triplet in enumerate(triplets):
        paths = [_proc_file(event.get("request_path")) for event in triplet]
        _require(paths == list(PROC_FILES),
                 f"procfs marker triplet {ordinal} has paths {paths!r}")
        for event in triplet:
            _require(event["context_status"]["truncated_context"] is False,
                     "procfs marker stack is frame-limit truncated")
        snapshots.append({
            "snapshot_index": len(snapshots),
            "event_indices": [event["index"] for event in triplet],
            "paths": paths,
        })

    def pair_windows(start_snapshot: int, count: int, kind: str) -> list[dict[str, Any]]:
        result: list[dict[str, Any]] = []
        for pair_index in range(count):
            before = snapshots[start_snapshot + 2 * pair_index]
            after = snapshots[start_snapshot + 2 * pair_index + 1]
            start_index = before["event_indices"][2]       # status open, inclusive
            end_index = after["event_indices"][0]           # io open, inclusive
            _require(start_index < end_index,
                     f"{kind} pair {pair_index} has inverted marker window")
            inside = [event for event in events
                      if start_index <= event["index"] <= end_index]
            mapping = [event for event in inside
                       if event["syscall"] in MAPPING_SYSCALLS
                       and event["classification"] == "phase"]
            result.append({
                "pair_index": pair_index,
                "before_snapshot_index": before["snapshot_index"],
                "after_snapshot_index": after["snapshot_index"],
                "start_event_index": start_index,
                "end_event_index": end_index,
                "traced_event_indices": [event["index"] for event in inside],
                "phase_mapping_event_indices": [event["index"] for event in mapping],
                "phase_mapping_syscalls": [event["syscall"] for event in mapping],
            })
        return result

    control_windows = pair_windows(0, CONTROL_PAIRS, "control")
    warmup_start = 2 * CONTROL_PAIRS
    warmup_windows = pair_windows(warmup_start, warmup, "warmup")
    measured_start = warmup_start + 2 * warmup
    measured_windows = pair_windows(measured_start, samples, "measured")

    sample_deltas, result_report = _raw_sample_deltas(report)
    _require(len(sample_deltas) == samples,
             f"raw sample_deltas count {len(sample_deltas)} differs from measured count {samples}")
    operation_metrics = result_report.get("operation_metrics")
    if isinstance(operation_metrics, Mapping) and "sample_indices" in operation_metrics:
        indices = operation_metrics["sample_indices"]
        _require(isinstance(indices, list) and len(indices) == samples,
                 "operation sample_indices are malformed")
        _require(sorted(indices) == list(range(samples)),
                 "operation sample_indices are not an acquisition-order permutation")
        elapsed_position = {acquisition: position for position, acquisition in enumerate(indices)}
    else:
        indices = list(range(samples))
        elapsed_position = {index: index for index in range(samples)}

    buckets: dict[str, dict[str, Any]] = {
        "zero": {"sample_indices": [], "phase_mapping_event_indices": [],
                 "minor_fault_values": {}},
        "78": {"sample_indices": [], "phase_mapping_event_indices": [],
               "minor_fault_values": {}},
        "other": {"sample_indices": [], "phase_mapping_event_indices": [],
                  "minor_fault_values": {}},
        "unavailable": {"sample_indices": [], "phase_mapping_event_indices": [],
                        "minor_fault_values": {}},
    }
    for index, (window, delta) in enumerate(zip(measured_windows, sample_deltas)):
        if delta is None:
            bucket = "unavailable"
            faults = None
        else:
            _require(isinstance(delta, Mapping), f"sample delta {index} is malformed")
            faults = delta.get("minor_faults")
            _require(isinstance(faults, int) and faults >= 0,
                     f"sample delta {index} minor_faults is malformed")
            bucket = "zero" if faults == 0 else "78" if faults == 78 else "other"
        window["sample_index"] = index
        window["acquisition_index"] = index
        window["report_sample_index"] = index
        window["elapsed_sorted_position"] = elapsed_position[index]
        window["minor_faults"] = faults
        window["fault_bucket"] = bucket
        row = buckets[bucket]
        row["sample_indices"].append(index)
        row["phase_mapping_event_indices"].extend(window["phase_mapping_event_indices"])
        if faults is not None:
            key = str(faults)
            row["minor_fault_values"][key] = row["minor_fault_values"].get(key, 0) + 1
    for row in buckets.values():
        row["sample_count"] = len(row["sample_indices"])
        row["phase_mapping_event_count"] = len(row["phase_mapping_event_indices"])
        row["sample_indices"] = list(row["sample_indices"])

    stack_resolution = "resolved" if unresolved_candidates == 0 else "unresolved"
    return {
        "status": "validated" if stack_resolution == "resolved" else "validated_unresolved_owner_stack",
        "stack_resolution": stack_resolution,
        "unresolved_marker_stack_count": unresolved_candidates,
        "extra_procfs_event_indices": [event["index"] for event in extras],
        "extra_procfs_event_count": len(extras),
        "owner_attribution_status": (
            "available" if stack_resolution == "resolved"
            else "unavailable_due_to_unresolved_marker_stacks"
        ),
        "expected_snapshot_count": expected_snapshots,
        "actual_snapshot_count": len(snapshots),
        "expected_marker_event_count": expected_marker_events,
        "actual_marker_event_count": len(marker_events),
        "control_pair_count": CONTROL_PAIRS,
        "warmup_pair_count": warmup,
        "measured_pair_count": samples,
        "snapshot_triplets": snapshots,
        "control_windows": control_windows,
        "warmup_windows": warmup_windows,
        "measured_windows": measured_windows,
        "fault_buckets": buckets,
        "window_scope": (
            "inclusive event interval from before /proc/self/status open through "
            "after /proc/self/io open; includes the tail of the before probe"
        ),
        "alignment": "last measured windows align to raw process_probe.sample_deltas in acquisition order; warmups remain separate",
    }


def parse_mapping(path: Path, plan: Mapping[str, Any], report: Any = None,
                  symbolization: Mapping[str, Any] | None = None) -> dict[str, Any]:
    """Parse one retained strace output and return a JSON-compatible summary.

    ``report`` is optional for callers that only need syscall attribution.  If
    supplied, it must be the raw ordinary-save report (or a path to it); the
    parser then validates all procfs marker triplets and attaches the measured
    fault buckets to their bounded mapping windows.  ``symbolization`` is an
    optional offline witness with ``binary.path`` and an exact ``frames`` map;
    when present it resolves blank local executable frames without changing the
    retained raw frame text.
    """

    _require(isinstance(plan, Mapping), "mapping plan is not an object")
    path = Path(path)
    text = _read_text(path)
    settings = _mapping_settings(plan)
    options = _options(settings)
    syscall_names = _trace_syscalls(plan)
    allowed = set(syscall_names)
    frame_limit = _frame_limit(options, DEFAULT_FRAME_LIMIT)
    owner = _owner_frames(plan)

    events: list[dict[str, Any]] = []
    current: dict[str, Any] | None = None
    terminal_code: int | None = None
    terminal_line: int | None = None
    unparsed_lines: list[dict[str, Any]] = []
    orphan_frames = 0
    trailing_lines = 0

    def finish() -> None:
        nonlocal current
        if current is None:
            return
        current["index"] = len(events)
        events.append(current)
        current = None

    for line_number, line in enumerate(text.splitlines(), start=1):
        if not line.strip():
            continue
        code = _terminal(line)
        if code is not None:
            finish()
            _require(terminal_code is None, "strace contains multiple tool terminals")
            terminal_code = code
            terminal_line = line_number
            continue
        if _is_killed_terminal(line):
            finish()
            raise AttributionError(f"strace tool was killed at line {line_number}")
        if terminal_code is not None:
            trailing_lines += 1
            continue
        frame = FRAME_RE.match(line)
        if frame is not None:
            if current is None:
                orphan_frames += 1
            else:
                current["stack"].append(frame.group("frame").strip())
            continue
        parsed = _parse_syscall(line, allowed)
        if parsed is not None:
            finish()
            current = parsed
            continue
        unparsed_lines.append({"line": line_number, "text": line})

    finish()
    _require(terminal_code is not None,
             "strace output does not contain +++ exited with N +++ terminal")
    _require(terminal_code == 0,
             f"strace tool terminal exited with {terminal_code}, expected zero")
    _require(not trailing_lines, "strace output has nonempty lines after tool terminal")

    for event in events:
        event["frame_limit"] = frame_limit
        status = _context_status(event["stack"], frame_limit)
        event["context_status"] = status
        event["unknown_context"] = status["unknown_context"]
        event["truncated_context"] = status["truncated_context"]

    symbolization_summary = _resolve_mapping_stacks(events, symbolization, frame_limit)
    for event in events:
        event["classification"] = _classify(event["classification_stack"], owner)

    groups = {
        group: _mapping_group(group, events, syscall_names, frame_limit, owner)
        for group in GROUPS
    }
    total_status = {
        "unknown_context_count": sum(row["unknown_context_count"] for row in groups.values()),
        "truncated_context_count": sum(row["truncated_context_count"] for row in groups.values()),
    }
    result: dict[str, Any] = {
        "schema_version": SCHEMA_VERSION,
        "kind": "strace_mapping",
        "source": {"path": str(path), "line_count": len(text.splitlines())},
        "tool": {
            "name": "strace",
            "successful_terminal": True,
            "terminal": {"line": terminal_line, "exit_code": terminal_code},
            "trace_syscalls": list(syscall_names),
            "stack_trace_frame_limit": frame_limit,
        },
        "symbolization": symbolization_summary,
        "owner_rule": {
            "phase": f"stack contains both {owner[0]!r} and {owner[1]!r}",
            "setup_publication": f"stack contains {owner[0]!r} without {owner[1]!r}",
            "other_unresolved": "all other stacks, including missing/unknown/truncated context",
            "owner_absence_inference": "forbidden when context is missing, unknown, or truncated",
        },
        "totals": _totals(events, syscall_names),
        "groups": groups,
        "phase_event_indices": groups["phase"]["event_indices"],
        "phase_stacks": [
            [frame["raw"] for frame in event["stack_evidence"]]
            for event in groups["phase"]["events"]
        ],
        "unknown_context_count": total_status["unknown_context_count"],
        "truncated_context_count": total_status["truncated_context_count"],
        "orphan_frame_count": orphan_frames,
        "unparsed_line_count": len(unparsed_lines),
        "unparsed_lines": unparsed_lines,
        "all_events_retained": False,
        "event_record_scope": "all traced syscall metadata and per-group stacks are retained; raw lines remain in the sibling strace artifact",
    }
    if report is not None:
        result["marker_windows"] = _marker_windows(events, plan, report, frame_limit, owner)
    else:
        result["marker_windows"] = {"status": "not_requested"}
    return result


def _nonnegative_counter(value: Any, label: str) -> int:
    _require(isinstance(value, int) and not isinstance(value, bool) and value >= 0,
             f"{label} must be a nonnegative integer")
    return value


def _dhat_header(data: Mapping[str, Any]) -> dict[str, Any]:
    _require(data.get("dhatFileVersion") == 2, "DHAT file version must be 2")
    _require(isinstance(data.get("pps"), list), "DHAT pps is not a list")
    _require(isinstance(data.get("ftbl"), list), "DHAT ftbl is not a list")
    _require(data.get("mode") == "heap", "DHAT mode is not heap")
    _require(isinstance(data.get("ftbl"), list)
             and all(isinstance(frame, str) for frame in data["ftbl"]),
             "DHAT frame table contains a non-string frame")
    for name in ("te", "tg"):
        if name in data:
            _nonnegative_counter(data[name], f"DHAT header {name}")
    return {
        "dhatFileVersion": data["dhatFileVersion"],
        "mode": data.get("mode"),
        "verb": data.get("verb"),
        "bklt": data.get("bklt"),
        "bkacc": data.get("bkacc"),
        "tu": data.get("tu"),
        "Mtu": data.get("Mtu"),
        "frame_count": len(data["ftbl"]),
        "pp_count": len(data["pps"]),
    }


ALLOC_COUNTERS = ("tb", "tbk", "tl", "mb", "mbk", "gb", "gbk", "eb", "ebk", "rb", "wb")


def _dhat_pp(pp: Mapping[str, Any], index: int, frame_count: int) -> dict[str, Any]:
    _require(isinstance(pp, Mapping), f"DHAT PP {index} is not an object")
    _require(isinstance(pp.get("fs"), list), f"DHAT PP {index} frame sequence is malformed")
    frame_indices: list[int] = []
    for frame_index in pp["fs"]:
        _require(isinstance(frame_index, int) and not isinstance(frame_index, bool),
                 f"DHAT PP {index} frame index is not an integer")
        _require(0 <= frame_index < frame_count,
                 f"DHAT PP {index} frame index {frame_index} is out of range")
        frame_indices.append(frame_index)
    for name in ("tb", "tbk"):
        _nonnegative_counter(pp.get(name), f"DHAT PP {index} {name}")
    for name in ALLOC_COUNTERS:
        if name in pp:
            _nonnegative_counter(pp[name], f"DHAT PP {index} {name}")
    if "acc" in pp:
        _require(isinstance(pp["acc"], list)
                 and all(isinstance(value, int) and not isinstance(value, bool)
                         for value in pp["acc"]),
                 f"DHAT PP {index} access counters are malformed")
    return {"frame_indices": frame_indices,
            "allocated_bytes": int(pp["tb"]),
            "allocated_blocks": int(pp["tbk"])}


def _allocation_group(group: str, contexts: Sequence[Mapping[str, Any]]) -> dict[str, Any]:
    selected = [context for context in contexts if context["classification"] == group]
    return {
        "classification": group,
        "pp_indices": [context["pp_index"] for context in selected],
        "contexts": [dict(context) for context in selected],
        "totals": {
            "pp_count": len(selected),
            "allocated_bytes": sum(context["allocated_bytes"] for context in selected),
            "allocated_blocks": sum(context["allocated_blocks"] for context in selected),
        },
        "owner_absence_inference": "forbidden_when_context_is_missing_unknown_or_truncated",
    }


def parse_allocation(path: Path, plan: Mapping[str, Any]) -> dict[str, Any]:
    """Validate one DHAT v2 file and aggregate cumulative request totals."""

    _require(isinstance(plan, Mapping), "allocation plan is not an object")
    path = Path(path)
    data = _read_json(path)
    _require(isinstance(data, Mapping), "DHAT file is not an object")
    settings = _allocation_settings(plan)
    options = _options(settings)
    frame_limit = _frame_limit(options, DEFAULT_DHAT_FRAME_LIMIT)
    owner = _owner_frames(plan)
    header = _dhat_header(data)
    frame_table = data["ftbl"]

    contexts: list[dict[str, Any]] = []
    unknown_context_count = 0
    truncated_context_count = 0
    unknown_frame_count = 0
    for index, pp in enumerate(data["pps"]):
        raw = _dhat_pp(pp, index, len(frame_table))
        frames = [frame_table[frame_index] for frame_index in raw["frame_indices"]]
        status = _context_status(frames, frame_limit)
        if status["unknown_context"]:
            unknown_context_count += 1
        if status["truncated_context"]:
            truncated_context_count += 1
        unknown_frame_count += status["unknown_frame_count"]
        contexts.append({
            "pp_index": index,
            "frame_indices": raw["frame_indices"],
            "frames": frames,
            "allocated_bytes": raw["allocated_bytes"],
            "allocated_blocks": raw["allocated_blocks"],
            "classification": _classify(frames, owner),
            "owner_evidence": _owner_evidence(frames, status, owner),
            "unknown_context": status["unknown_context"],
            "truncated_context": status["truncated_context"],
        })

    groups = {
        group: _allocation_group(group, contexts)
        for group in GROUPS
    }
    totals = {
        "pp_count": len(contexts),
        "allocated_bytes": sum(context["allocated_bytes"] for context in contexts),
        "allocated_blocks": sum(context["allocated_blocks"] for context in contexts),
    }
    result: dict[str, Any] = {
        "schema_version": SCHEMA_VERSION,
        "kind": "dhat_allocation",
        "source": {"path": str(path)},
        "tool": {
            "name": "valgrind-dhat",
            "validated_file_version": 2,
            "num_callers": frame_limit,
        },
        "header": header,
        "owner_rule": {
            "phase": f"frame sequence contains both {owner[0]!r} and {owner[1]!r}",
            "setup_publication": f"frame sequence contains {owner[0]!r} without {owner[1]!r}",
            "other_unresolved": "all other frame sequences, including missing/unknown/truncated context",
            "owner_absence_inference": "forbidden when context is missing, unknown, or truncated",
        },
        "totals": totals,
        "cumulative_allocated": {
            "bytes": totals["allocated_bytes"],
            "blocks": totals["allocated_blocks"],
            "interpretation": "sum of DHAT tb/tbk allocation requests across PP records",
        },
        "groups": groups,
        "phase_pp_indices": groups["phase"]["pp_indices"],
        "phase_contexts": groups["phase"]["contexts"],
        "unknown_context_count": unknown_context_count,
        "truncated_context_count": truncated_context_count,
        "unknown_frame_count": unknown_frame_count,
        "frame_limit_caveat": (
            "A frame sequence reaching the configured --num-callers limit may be "
            "truncated; missing owner markers therefore do not establish owner absence."
        ),
        "excluded_from_totals": [
            "mb", "mbk", "gb", "gbk", "eb", "ebk",
            "native allocator behavior", "peak values as additive quantities",
        ],
    }
    return result


__all__ = ["AttributionError", "parse_mapping", "parse_allocation"]
