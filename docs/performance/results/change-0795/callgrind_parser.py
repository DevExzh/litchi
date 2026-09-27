#!/usr/bin/env python3
"""A small, strict Callgrind parser used by the 0795 offline analysis.

The older 0784 reader intentionally knew about one event (``Ir``).  This
reader keeps the same useful checks around compressed function names and
call-graph records, while representing every cost as a vector.  It is
deliberately independent of Valgrind and has no benchmark or build side
effects; it only reads a retained Callgrind text file.
"""

from __future__ import annotations

import collections
import hashlib
import re
from pathlib import Path
from typing import Any, Iterable, Mapping, NoReturn, Sequence


DEFAULT_EVENTS = ("Ir", "Bc", "Bcm", "Bi", "Bim")
_INTEGER = re.compile(r"^[+-]?\d[\d,]*$")
_FUNCTION = re.compile(r"^(fn|cfn)=\((\d+)\)(?:\s+(.*))?$")
_CALLS = re.compile(r"^calls=([+-]?\d[\d,]*)(?:\s+(.*))?$")
_POSITION = re.compile(r"^(?:[+-]?\d+|\*)$")


class CallgrindError(ValueError):
    """A malformed or internally contradictory Callgrind artifact."""


def _fail(message: str) -> NoReturn:
    raise CallgrindError(message)


def _require(condition: bool, message: str) -> None:
    if not condition:
        _fail(message)


def _integer(token: str, label: str) -> int:
    if not _INTEGER.fullmatch(token):
        _fail(f"{label}: expected an integer, got {token!r}")
    try:
        return int(token.replace(",", ""))
    except ValueError as error:  # pragma: no cover - guarded by the regexp
        raise CallgrindError(f"{label}: invalid integer {token!r}") from error


def vector_zero(events: Sequence[str]) -> tuple[int, ...]:
    return (0,) * len(events)


def vector_add(left: Sequence[int], right: Sequence[int]) -> tuple[int, ...]:
    _require(len(left) == len(right), "cannot add vectors with different event counts")
    return tuple(a + b for a, b in zip(left, right))


def vector_sum(vectors: Iterable[Sequence[int]], events: Sequence[str]) -> tuple[int, ...]:
    result = vector_zero(events)
    for vector in vectors:
        result = vector_add(result, vector)
    return result


def vector_equal(left: Sequence[int] | None, right: Sequence[int] | None) -> bool:
    return left is not None and right is not None and tuple(left) == tuple(right)


def vector_mapping(values: Sequence[int] | None, events: Sequence[str]) -> dict[str, int] | None:
    if values is None:
        return None
    _require(len(values) == len(events), "vector/event cardinality differs")
    return {event: int(value) for event, value in zip(events, values)}


def _parse_vector(text: str, count: int, label: str) -> tuple[int, ...]:
    fields = text.split()
    # Summary/totals lines follow the same Callgrind shorthand as cost lines:
    # trailing zero events may be omitted (a termination part commonly says
    # simply ``summary: 0`` even when branch simulation enabled five events).
    _require(0 < len(fields) <= count, f"{label}: expected at most {count} event values")
    values = [_integer(token, f"{label} event {index}")
              for index, token in enumerate(fields)]
    values.extend([0] * (count - len(values)))
    return tuple(values)


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _name(names: Mapping[int, str], function_id: int) -> str:
    return names.get(function_id, "") or "<unnamed>"


def _add_name(names: dict[int, str], function_id: int, name: str, label: str) -> None:
    if not name:
        return
    previous = names.get(function_id)
    _require(previous is None or previous == name,
             f"{label}: function id {function_id} has conflicting names")
    names[function_id] = name


def _header(lines: Sequence[str], expected_events: Sequence[str] | None) -> dict[str, Any]:
    events: list[str] | None = None
    summary: tuple[int, ...] | None = None
    totals: tuple[int, ...] | None = None
    part: int | None = None
    trigger: str | None = None
    command: str | None = None
    version: str | None = None
    positions: str | None = None

    for line_number, raw in enumerate(lines, 1):
        line = raw.strip()
        if line.startswith("events:"):
            _require(events is None, f"line {line_number}: duplicate events header")
            fields = line.split(":", 1)[1].split()
            _require(fields and len(fields) == len(set(fields)),
                     f"line {line_number}: events header is empty or duplicated")
            events = fields
        elif line.startswith("summary:"):
            _require(events is not None,
                     f"line {line_number}: summary precedes events header")
            _require(summary is None, f"line {line_number}: duplicate summary")
            summary = _parse_vector(line.split(":", 1)[1], len(events), "summary")
        elif line.startswith("totals:"):
            _require(events is not None,
                     f"line {line_number}: totals precedes events header")
            _require(totals is None, f"line {line_number}: duplicate totals")
            totals = _parse_vector(line.split(":", 1)[1], len(events), "totals")
        elif line.startswith("part:"):
            _require(part is None, f"line {line_number}: duplicate part")
            part = _integer(line.split(":", 1)[1].strip(), "part")
        elif line.startswith("positions:"):
            _require(positions is None, f"line {line_number}: duplicate positions")
            positions = line.split(":", 1)[1].strip()
        elif line.startswith("desc: Trigger:"):
            _require(trigger is None, f"line {line_number}: duplicate trigger")
            trigger = line.split("Trigger:", 1)[1].strip()
        elif line.startswith("cmd:") and command is None:
            command = line.split(":", 1)[1].strip()
        elif line.startswith("version:") and version is None:
            version = line.split(":", 1)[1].strip()

    _require(events is not None, "Callgrind events header is missing")
    if expected_events is not None:
        _require(tuple(events) == tuple(expected_events),
                 f"Callgrind event set differs: {events!r} != {list(expected_events)!r}")
    _require(summary is not None, "Callgrind summary is missing")
    return {
        "events": tuple(events),
        "summary": summary,
        "totals": totals,
        "part": part,
        "trigger": trigger,
        "command": command,
        "version": version,
        "positions": positions,
    }


def _cost_line(line: str, event_count: int) -> tuple[str, tuple[int, ...]] | None:
    fields = line.split()
    # Callgrind permits trailing zero event values to be omitted.  For
    # example, with ``Ir Bc Bcm Bi Bim`` the line ``0 7`` means
    # ``7 0 0 0 0``.  It never omits a value in the middle of the vector.
    if len(fields) < 2 or not _POSITION.fullmatch(fields[0]):
        return None
    value_fields = fields[1:]
    if len(value_fields) > event_count:
        return None
    values: list[int] = []
    for index, token in enumerate(value_fields):
        if not _INTEGER.fullmatch(token):
            return None
        values.append(_integer(token, f"cost event {index}"))
    values.extend([0] * (event_count - len(values)))
    return fields[0], tuple(values)


def position_kind(position: str) -> str:
    if position == "*":
        return "wildcard"
    if position.startswith(("+", "-")):
        return "relative"
    return "absolute"


def parse_callgrind(
    path: str | Path,
    *,
    expected_events: Sequence[str] | None = DEFAULT_EVENTS,
    allow_empty: bool = False,
) -> dict[str, Any]:
    """Parse one Callgrind part and return a graph with vector-valued costs.

    The returned ``functions`` map uses integer function ids for convenient
    graph work.  Its ``self`` and edge ``cost`` values are tuples ordered by
    ``header['events']``.  The result is intentionally an in-memory analysis
    object; callers that serialize a report should convert those tuples with
    :func:`vector_mapping`.
    """

    source = Path(path)
    _require(source.is_file() and not source.is_symlink(),
             f"Callgrind artifact is not a regular file: {source}")
    try:
        text = source.read_text(encoding="utf-8", errors="replace")
    except OSError as error:
        raise CallgrindError(f"cannot read {source}: {error}") from error
    lines = text.splitlines()
    header = _header(lines, expected_events)
    events = tuple(header["events"])
    event_count = len(events)
    label = str(source)

    # Resolve compressed names in a first pass.  Callgrind may declare an id
    # with its name on one record and refer to it without a name later.
    names: dict[int, str] = {}
    declarations = 0
    for line_number, raw in enumerate(lines, 1):
        match = _FUNCTION.match(raw.strip())
        if match and match.group(3):
            declarations += 1
            _add_name(names, int(match.group(2)), match.group(3),
                      f"{label}:{line_number}")

    functions: dict[int, dict[str, Any]] = {}
    current_id: int | None = None
    pending: dict[str, Any] | None = None
    cost_records = 0
    edge_records = 0
    positions = collections.Counter[str]()

    for line_number, raw in enumerate(lines, 1):
        line = raw.strip()
        function_match = _FUNCTION.match(line)
        if function_match:
            kind, text_id, declared_name = function_match.groups()
            function_id = int(text_id)
            if declared_name:
                _add_name(names, function_id, declared_name,
                          f"{label}:{line_number}")
            if kind == "fn":
                _require(pending is None or pending.get("calls") is None,
                         f"{label}:{line_number}: cfn has no cost record")
                current_id = function_id
                pending = None
                function = functions.setdefault(
                    function_id,
                    {
                        "id": function_id,
                        "records": 0,
                        "self": vector_zero(events),
                        "edges": [],
                    },
                )
                function["records"] += 1
            else:
                _require(current_id is not None,
                         f"{label}:{line_number}: cfn appears outside fn")
                _require(pending is None,
                         f"{label}:{line_number}: cfn replaced an unresolved edge")
                pending = {
                    "callee_id": function_id,
                    "calls": None,
                    "position": None,
                    "line": line_number,
                }
            continue

        calls_match = _CALLS.match(line)
        if calls_match:
            _require(current_id is not None and pending is not None,
                     f"{label}:{line_number}: calls appears outside cfn")
            _require(pending["calls"] is None,
                     f"{label}:{line_number}: duplicate calls record")
            pending["calls"] = _integer(calls_match.group(1),
                                         f"{label}:{line_number} calls")
            trailing = (calls_match.group(2) or "").split()
            if trailing and _POSITION.fullmatch(trailing[0]):
                pending["position"] = trailing[0]
            continue

        if current_id is None:
            continue
        parsed = _cost_line(line, event_count)
        if parsed is None:
            continue
        position, cost = parsed
        positions[position_kind(position)] += 1
        cost_records += 1
        function = functions[current_id]
        if pending is not None:
            _require(pending["calls"] is not None,
                     f"{label}:{line_number}: cfn has a cost before calls")
            function["edges"].append(
                {
                    "caller_id": current_id,
                    "callee_id": pending["callee_id"],
                    "calls": pending["calls"],
                    "position": pending["position"] or position,
                    "cost": cost,
                    "line": pending["line"],
                }
            )
            edge_records += 1
            pending = None
        else:
            function["self"] = vector_add(function["self"], cost)

    _require(pending is None or pending.get("calls") is None,
             f"{label}: cfn has no cost record")
    for function_id, function in functions.items():
        function["name"] = names.get(function_id, "")

    summary = header["summary"]
    totals = header["totals"]
    self_total = vector_sum((function["self"] for function in functions.values()), events)
    empty_part = (summary == vector_zero(events) and not names and not functions
                  and cost_records == 0)
    if empty_part:
        _require(allow_empty, f"{label}: empty Callgrind part requires allow_empty")
    else:
        _require(cost_records > 0 or summary == vector_zero(events),
                 f"{label}: no cost records were parsed")
        _require(names or summary == vector_zero(events),
                 f"{label}: no compressed function names were resolved")

    return {
        "path": label,
        "sha256": _sha256(source),
        "bytes": source.stat().st_size,
        "header": header,
        "names": names,
        "functions": functions,
        "self_total": self_total,
        "statistics": {
            "function_records": sum(function["records"] for function in functions.values()),
            "unique_functions": len(functions),
            "compressed_name_declarations": declarations,
            "cost_records": cost_records,
            "call_edge_records": edge_records,
            "position_kinds": dict(sorted(positions.items())),
        },
    }


def aggregate_edges(
    parsed: Mapping[str, Any], function_id: int,
) -> list[dict[str, Any]]:
    """Group a function's direct call edges by callee id.

    All event vectors are retained, including edges with zero Ir but nonzero
    branch counters.  This matters when the selected event is a branch event
    and is also why the old Ir-only ``> 0`` filter is intentionally absent.
    """

    events = tuple(parsed["header"]["events"])
    names = parsed["names"]
    function = parsed["functions"].get(function_id)
    if function is None:
        return []
    grouped: dict[int, dict[str, Any]] = {}
    for edge in function["edges"]:
        callee_id = int(edge["callee_id"])
        item = grouped.setdefault(
            callee_id,
            {
                "callee_id": callee_id,
                "callee": _name(names, callee_id),
                "calls": 0,
                "edge_count": 0,
                "cost": vector_zero(events),
                "positions": [],
            },
        )
        item["calls"] += int(edge["calls"])
        item["edge_count"] += 1
        item["cost"] = vector_add(item["cost"], edge["cost"])
        item["positions"].append(edge["position"])
    return sorted(
        grouped.values(),
        key=lambda item: (-item["cost"][0], -max(item["cost"]), item["callee"], item["callee_id"]),
    )


def incoming_edges(parsed: Mapping[str, Any], target_id: int) -> list[dict[str, Any]]:
    """Return every positive or counter-bearing incoming edge to a function."""

    names = parsed["names"]
    result: list[dict[str, Any]] = []
    for function in parsed["functions"].values():
        for edge in function["edges"]:
            if edge["callee_id"] != target_id:
                continue
            if edge["calls"] == 0 and not any(edge["cost"]):
                continue
            result.append(
                {
                    **edge,
                    "caller": _name(names, edge["caller_id"]),
                    "callee": _name(names, edge["callee_id"]),
                }
            )
    return sorted(
        result,
        key=lambda item: (-item["cost"][0], item["caller"], item["caller_id"], item["line"]),
    )


def function_inclusive(parsed: Mapping[str, Any], function_id: int) -> tuple[int, ...]:
    """Sum incoming edge costs for a diagnostic inclusive value.

    A function can have several callers, so this value can overlap.  Reports
    must label it as such and must never use it as a disjoint partition.
    """

    events = tuple(parsed["header"]["events"])
    return vector_sum((edge["cost"] for edge in incoming_edges(parsed, function_id)), events)


def function_calls_in(parsed: Mapping[str, Any], function_id: int) -> int:
    return sum(int(edge["calls"]) for edge in incoming_edges(parsed, function_id))


def function_calls_out(parsed: Mapping[str, Any], function_id: int) -> int:
    function = parsed["functions"].get(function_id)
    return 0 if function is None else sum(int(edge["calls"]) for edge in function["edges"])


def function_view(parsed: Mapping[str, Any], function_id: int) -> dict[str, Any]:
    events = tuple(parsed["header"]["events"])
    function = parsed["functions"][function_id]
    children = aggregate_edges(parsed, function_id)
    return {
        "id": function_id,
        "name": _name(parsed["names"], function_id),
        "records": function["records"],
        "self": function["self"],
        "inclusive": function_inclusive(parsed, function_id),
        "calls_in": function_calls_in(parsed, function_id),
        "calls_out": function_calls_out(parsed, function_id),
        "direct_children": children,
        "direct_children_cost": vector_sum((child["cost"] for child in children), events),
    }


def serializable_edge(edge: Mapping[str, Any], events: Sequence[str]) -> dict[str, Any]:
    """Convert an internal edge to stable report data."""

    return {
        "caller_id": int(edge["caller_id"]),
        "caller": str(edge.get("caller", "<unnamed>")),
        "callee_id": int(edge["callee_id"]),
        "callee": str(edge.get("callee", "<unnamed>")),
        "calls": int(edge["calls"]),
        "position": edge.get("position"),
        "cost": vector_mapping(edge["cost"], events),
    }


def serializable_child(child: Mapping[str, Any], events: Sequence[str]) -> dict[str, Any]:
    return {
        "callee_id": int(child["callee_id"]),
        "callee": str(child["callee"]),
        "calls": int(child["calls"]),
        "edge_count": int(child["edge_count"]),
        "cost": vector_mapping(child["cost"], events),
        "positions": list(child.get("positions", [])),
    }
