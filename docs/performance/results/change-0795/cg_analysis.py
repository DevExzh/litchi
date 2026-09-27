#!/usr/bin/env python3
"""Offline multi event Callgrind analysis for performance packet 0795.

The root agent owns all builds and Callgrind children.  This module runs after
those children have terminated and reads only the retained receipts, JSON
reports, positive numbered dumps, and zero termination dumps.  It deliberately
keeps owner qualification as a reported fact: if the exact owner cannot be
resolved, the profile remains in the report with ``owner_qualified: false``
and the reason is retained.
"""

from __future__ import annotations

import argparse
import collections
import hashlib
import json
import re
import sys
from pathlib import Path
from typing import Any, Iterable, Mapping, NoReturn, Sequence

import custody as c
from callgrind_parser import (
    DEFAULT_EVENTS,
    CallgrindError,
    aggregate_edges,
    function_inclusive,
    function_view,
    incoming_edges,
    parse_callgrind,
    serializable_child,
    serializable_edge,
    vector_add,
    vector_equal,
    vector_mapping,
    vector_sum,
    vector_zero,
)


HERE = Path(__file__).resolve().parent
EVENTS = tuple(DEFAULT_EVENTS)
OWNER = "namespace_uri_probe::capture_region_0793"
SHAPES = ("tiny", "medium", "large")
LEGS = ("before", "after")
ANALYSIS_JSON = HERE / "callgrind-analysis.json"
ANALYSIS_MD = HERE / "callgrind-analysis.md"

SELECTED_PATTERNS: dict[str, re.Pattern[str]] = {
    "notes": re.compile(r"notes", re.IGNORECASE),
    "inspect": re.compile(r"inspect", re.IGNORECASE),
    "checked_attributes": re.compile(r"CheckedAttributes"),
    "iter_state": re.compile(r"IterState"),
    "allocation_named_functions": re.compile(
        r"(?:__rust_alloc|__rust_dealloc|__rust_realloc|alloc(?:ate|ation)?|calloc|malloc|realloc|dealloc)",
        re.IGNORECASE,
    ),
    "memcpy": re.compile(r"(?:memcpy|memmove|memset)", re.IGNORECASE),
    "opened_presentation": re.compile(r"opened_presentation", re.IGNORECASE),
}


class EvidenceError(ValueError):
    """A retained artifact is missing, malformed, or contradictory."""


def fail(message: str) -> NoReturn:
    raise EvidenceError(message)


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def read_json(path: Path) -> Any:
    require(path.is_file() and not path.is_symlink(), f"missing JSON: {path}")
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        fail(f"invalid JSON {path}: {error}")


def relative(path: Path) -> str:
    try:
        return str(path.resolve().relative_to(HERE))
    except ValueError as error:
        raise EvidenceError(f"path escapes packet: {path}") from error


def packet_path(raw: Any, label: str) -> Path:
    require(isinstance(raw, str) and raw, f"{label}: path is missing")
    marker = f"/{HERE.name}/"
    if marker in raw:
        path = HERE / raw.split(marker, 1)[1]
    else:
        candidate = Path(raw)
        path = candidate if candidate.is_absolute() else HERE / candidate
    path = path.resolve()
    require(path == HERE or HERE in path.parents,
            f"{label}: path escapes packet: {raw}")
    return path


def external_path(raw: Any, label: str) -> Path | None:
    if not isinstance(raw, str) or not raw:
        return None
    marker = f"/{HERE.name}/"
    if marker in raw:
        return (HERE / raw.split(marker, 1)[1]).resolve()
    path = Path(raw)
    return path.resolve() if path.is_absolute() else None


def artifact(value: Any, label: str) -> dict[str, Any]:
    require(isinstance(value, dict), f"{label}: artifact is not an object")
    path = packet_path(value.get("path"), label)
    expected_sha = value.get("sha256")
    expected_bytes = value.get("bytes")
    require(isinstance(expected_sha, str) and re.fullmatch(r"[0-9a-f]{64}", expected_sha),
            f"{label}: invalid SHA-256")
    require(type(expected_bytes) is int and expected_bytes >= 0,
            f"{label}: invalid byte count")
    require(path.is_file() and not path.is_symlink(), f"{label}: artifact is missing: {path}")
    require(path.stat().st_size == expected_bytes,
            f"{label}: artifact byte count changed")
    actual_sha = c.sha(path)
    require(actual_sha == expected_sha, f"{label}: artifact SHA-256 changed")
    return {"path": relative(path), "sha256": expected_sha, "bytes": expected_bytes}


def vector_nonzero(values: Sequence[int] | None) -> bool:
    return values is not None and any(int(value) != 0 for value in values)


def plan_path() -> Path:
    frozen = HERE / "frozenplan.json"
    return frozen if frozen.is_file() else HERE / "plan.json"


def validate_plan() -> tuple[dict[str, Any], Path]:
    path = plan_path()
    plan = read_json(path)
    require(isinstance(plan, dict), "frozen plan is not an object")
    require(plan.get("schema") == "litchi.performance.0795.v1", "plan schema changed")
    require(plan.get("cpu") == 12, "Callgrind CPU changed")
    require(plan.get("owner") == OWNER, "Callgrind owner changed")
    require(plan.get("source_allowlist") and isinstance(plan["source_allowlist"], list),
            "plan source allowlist is missing")
    callgrind = plan.get("callgrind")
    require(isinstance(callgrind, dict), "callgrind plan is missing")
    require(callgrind.get("branch_sim") is True, "branch simulation policy changed")
    require(callgrind.get("events") == list(EVENTS), "Callgrind event set changed")
    require(callgrind.get("samples") == 1 and callgrind.get("warmup") == 0,
            "Callgrind sample policy changed")
    require(callgrind.get("positive_dumps") == 1 and callgrind.get("termination_empty") is True,
            "Callgrind dump policy changed")
    shape_orders = callgrind.get("shape_orders")
    leg_orders = callgrind.get("leg_orders")
    require(shape_orders == [list(SHAPES), ["large", "medium", "tiny"]],
            "Callgrind shape order changed")
    require(leg_orders == [["before", "after"], ["after", "before"]],
            "Callgrind leg order changed")
    require(callgrind.get("shapes") == list(SHAPES), "Callgrind shape set changed")
    return plan, path


def expected_jobs(plan: Mapping[str, Any]) -> list[tuple[int, str, str]]:
    callgrind = plan["callgrind"]
    shapes = callgrind["shape_orders"]
    legs = callgrind["leg_orders"]
    require(len(shapes) == len(legs), "shape and leg order cardinalities differ")
    jobs: list[tuple[int, str, str]] = []
    for repeat, (shape_order, leg_order) in enumerate(zip(shapes, legs)):
        for shape in shape_order:
            for leg in leg_order:
                jobs.append((repeat, shape, leg))
    return jobs


def receipt_binary(value: Any, label: str) -> dict[str, Any] | str:
    """Validate the shape of a receipt binary while tolerating cleaned binaries."""

    if isinstance(value, str):
        require(value, f"{label}: binary path is empty")
        return value
    require(isinstance(value, dict), f"{label}: binary identity is missing")
    path = value.get("path")
    expected_sha = value.get("sha256")
    expected_bytes = value.get("bytes")
    require(isinstance(path, str) and path, f"{label}: binary path is missing")
    require(isinstance(expected_sha, str) and re.fullmatch(r"[0-9a-f]{64}", expected_sha),
            f"{label}: binary SHA-256 is invalid")
    require(type(expected_bytes) is int and expected_bytes > 0,
            f"{label}: binary byte count is invalid")
    live = external_path(path, label)
    if live is not None and live.is_file() and not live.is_symlink():
        require(live.stat().st_size == expected_bytes, f"{label}: binary byte count changed")
        require(c.sha(live) == expected_sha, f"{label}: binary SHA-256 changed")
    return dict(value)


def command_check(command: Any, stem: str, row: Mapping[str, Any]) -> list[str]:
    label = f"profiles/{stem}"
    require(isinstance(command, list) and all(isinstance(item, str) for item in command),
            f"{label}: command is not a string list")
    text = command
    require("--tool=callgrind" in text, f"{label}: command does not use Callgrind")
    require("--branch-sim=yes" in text, f"{label}: branch simulation flag is missing")
    require("--collect-atstart=no" in text, f"{label}: collection-at-start policy changed")
    require("--toggle-collect=" + OWNER in text, f"{label}: owner toggle is missing")
    require("--zero-before=" + OWNER in text, f"{label}: zero-before owner is missing")
    require("--dump-after=" + OWNER in text, f"{label}: dump-after owner is missing")
    raw_options = [item for item in text if item.startswith("--callgrind-out-file=")]
    require(len(raw_options) == 1, f"{label}: Callgrind output option is not unique")
    raw_path = packet_path(raw_options[0].split("=", 1)[1], f"{label} raw output")
    require(raw_path == HERE / "profiles" / f"{stem}.callgrind",
            f"{label}: raw output path changed")
    output_options = [item for item in text if item.startswith("--output=")]
    # The probe driver may use either --output=path or two argv entries.
    output_path: Path | None = None
    if len(output_options) == 1:
        output_path = packet_path(output_options[0].split("=", 1)[1], f"{label} JSON output")
    elif "--output" in text:
        index = text.index("--output")
        require(index + 1 < len(text), f"{label}: JSON output path is missing")
        output_path = packet_path(text[index + 1], f"{label} JSON output")
    require(output_path == HERE / "profiles" / f"{stem}.json",
            f"{label}: JSON output path changed")
    for flag, value in (("--shape", row["shape"]), ("--samples", "1"), ("--warmup", "0")):
        require(flag in text, f"{label}: {flag} is missing")
        index = text.index(flag)
        require(index + 1 < len(text) and text[index + 1] == value,
                f"{label}: {flag} value changed")
    return list(command)


def receipt_rows(plan: Mapping[str, Any]) -> tuple[list[dict[str, Any]], dict[str, Any]]:
    path = HERE / "profiles" / "receipts.json"
    receipts = read_json(path)
    require(isinstance(receipts, list), "profile receipts are not a list")
    jobs = expected_jobs(plan)
    require(len(receipts) == len(jobs),
            f"profile receipt count changed: {len(receipts)} != {len(jobs)}")
    retained: list[dict[str, Any]] = []
    for row, (repeat, shape, leg) in zip(receipts, jobs):
        label = f"profiles/{repeat}-{shape}-{leg}"
        require(isinstance(row, dict), f"{label}: receipt is not an object")
        require(row.get("repeat") == repeat and row.get("shape") == shape
                and row.get("leg") == leg, f"{label}: receipt matrix position changed")
        require(row.get("exit_code") == 0, f"{label}: profiler child failed")
        require(type(row.get("started")) in (int, float)
                and type(row.get("ended")) in (int, float)
                and row["ended"] >= row["started"],
                f"{label}: receipt timestamps are invalid")
        command = command_check(row.get("command"), f"{repeat}-{shape}-{leg}", row)
        receipt_binary(row.get("binary"), f"{label} binary")
        stem = f"{repeat}-{shape}-{leg}"
        artifacts = row.get("artifacts")
        require(isinstance(artifacts, dict), f"{label}: artifact map is missing")
        expected_names = {
            f"{stem}.json",
            f"{stem}.log",
            f"{stem}.callgrind",
            f"{stem}.callgrind.1",
        }
        require(set(artifacts) == expected_names,
                f"{label}: artifact set differs from {sorted(expected_names)}")
        checked_artifacts = {
            name: artifact(descriptor, f"{label} {name}")
            for name, descriptor in sorted(artifacts.items())
        }
        report_path = packet_path(artifacts[f"{stem}.json"]["path"], f"{label} report")
        report = read_json(report_path)
        require(isinstance(report, dict), f"{label}: probe report is not an object")
        retained.append(
            {
                "repeat": repeat,
                "shape": shape,
                "leg": leg,
                "stem": stem,
                "command": command,
                "binary": row["binary"],
                "started": row["started"],
                "ended": row["ended"],
                "artifacts": checked_artifacts,
                "report": report,
            }
        )
    return retained, {"path": relative(path), "sha256": c.sha(path), "count": len(retained)}


def serializable_vector(values: Sequence[int], events: Sequence[str] = EVENTS) -> dict[str, int]:
    result = vector_mapping(values, events)
    require(result is not None, "internal vector is missing")
    return result


def serializable_view(view: Mapping[str, Any], events: Sequence[str]) -> dict[str, Any]:
    return {
        "id": int(view["id"]),
        "name": str(view["name"]),
        "records": int(view["records"]),
        "self": serializable_vector(view["self"], events),
        "inclusive": serializable_vector(view["inclusive"], events),
        "calls_in": int(view["calls_in"]),
        "calls_out": int(view["calls_out"]),
        "direct_children": [serializable_child(child, events)
                             for child in view["direct_children"]],
        "direct_children_cost": serializable_vector(view["direct_children_cost"], events),
    }


def dominant_path(parsed: Mapping[str, Any], owner_id: int, depth: int = 12) -> list[dict[str, Any]]:
    """Follow the highest Ir direct edge for a readable nested path."""

    events = tuple(parsed["header"]["events"])
    current = owner_id
    seen: set[int] = set()
    path: list[dict[str, Any]] = []
    incoming_cost: Sequence[int] | None = None
    for _ in range(depth):
        if current in seen or current not in parsed["functions"]:
            break
        seen.add(current)
        view = function_view(parsed, current)
        node: dict[str, Any] = {
            "id": current,
            "name": view["name"],
            "self": serializable_vector(view["self"], events),
            "inclusive": (serializable_vector(incoming_cost, events)
                           if incoming_cost is not None else None),
        }
        children = [child for child in view["direct_children"]
                    if child["callee_id"] in parsed["functions"]]
        path.append(node)
        if not children:
            break
        child = children[0]
        node["next_edge"] = serializable_child(child, events)
        incoming_cost = child["cost"]
        current = child["callee_id"]
    return path


def opened_descendants(parsed: Mapping[str, Any], owner_id: int, depth: int = 64) -> dict[str, Any]:
    """Find public opened_presentation descendants without claiming fractions."""

    events = tuple(parsed["header"]["events"])
    matches: list[dict[str, Any]] = []
    queue: collections.deque[tuple[int, list[dict[str, Any]]]] = collections.deque([(owner_id, [])])
    seen: set[tuple[int, tuple[int, ...]]] = set()
    while queue and len(matches) < 64:
        current, prior_path = queue.popleft()
        if len(prior_path) >= depth:
            continue
        for child in aggregate_edges(parsed, current):
            child_id = child["callee_id"]
            if child_id not in parsed["functions"]:
                continue
            edge = serializable_child(child, events)
            next_path = prior_path + [edge]
            state = (child_id, tuple(item["callee_id"] for item in next_path[-8:]))
            if state in seen:
                continue
            seen.add(state)
            child_name = parsed["functions"][child_id]["name"] or "<unnamed>"
            if "opened_presentation" in child_name.lower() and child_id != owner_id:
                matches.append(
                    {
                        "id": child_id,
                        "name": child_name,
                        "path": next_path,
                        "self": serializable_vector(parsed["functions"][child_id]["self"], events),
                        "inclusive": serializable_vector(child["cost"], events),
                        "calls": int(child["calls"]),
                    }
                )
            queue.append((child_id, next_path))
    return {
        "matches": sorted(matches, key=lambda item: (item["name"], item["id"])),
        "qualified": bool(matches),
        "nested_costs_overlap": True,
        "fractions_reported": False,
    }


def owner_attribution(parsed: Mapping[str, Any]) -> dict[str, Any]:
    events = tuple(parsed["header"]["events"])
    functions = parsed["functions"]
    names = parsed["names"]
    owner_ids = sorted(function_id for function_id, function in functions.items()
                       if function.get("name") == OWNER)
    failures: list[str] = []
    selected: dict[str, Any] = {
        "name": OWNER,
        "ids": owner_ids,
        "owner_qualified": False,
        "qualification_failures": failures,
    }
    if len(owner_ids) != 1:
        failures.append(f"expected exactly one exact owner function, found {owner_ids}")
        return selected

    owner_id = owner_ids[0]
    incoming = incoming_edges(parsed, owner_id)
    positive_incoming = [edge for edge in incoming
                         if int(edge["calls"]) > 0 and vector_nonzero(edge["cost"])]
    if len(positive_incoming) != 1:
        failures.append(f"expected one positive owner incoming edge, found {len(positive_incoming)}")
    function = functions[owner_id]
    view = function_view(parsed, owner_id)
    direct_cost = view["direct_children_cost"]
    if positive_incoming:
        owner_edge = positive_incoming[0]
        if not vector_equal(owner_edge["cost"], parsed["header"]["summary"]):
            failures.append("positive owner incoming cost differs from Callgrind summary")
        if not vector_equal(vector_add(function["self"], direct_cost), owner_edge["cost"]):
            failures.append("owner self plus direct child inclusive cost does not partition owner")
    selected.update(
        {
            "id": owner_id,
            "view": serializable_view(view, events),
            "incoming": [serializable_edge(edge, events) for edge in incoming],
            "direct_children": [serializable_child(child, events)
                                for child in view["direct_children"]],
            "dominant_path": dominant_path(parsed, owner_id),
            "public_opened_presentation_descendant": opened_descendants(parsed, owner_id),
        }
    )
    if not failures:
        selected["owner_qualified"] = True
    return selected


def selected_row(parsed: Mapping[str, Any], function_id: int) -> dict[str, Any]:
    events = tuple(parsed["header"]["events"])
    view = function_view(parsed, function_id)
    incoming = incoming_edges(parsed, function_id)
    return {
        "id": function_id,
        "name": view["name"],
        "records": int(view["records"]),
        "self": serializable_vector(view["self"], events),
        # This is the sum of incoming edge costs.  It is intentionally marked
        # overlapping because a function can have multiple callers.
        "inclusive": serializable_vector(view["inclusive"], events),
        "calls_in": int(view["calls_in"]),
        "calls_out": int(view["calls_out"]),
        "incoming_edge_count": len(incoming),
        "inclusive_overlaps": len(incoming) > 1,
        "direct_children": [serializable_child(child, events)
                             for child in view["direct_children"]],
    }


def selected_symbols(parsed: Mapping[str, Any]) -> dict[str, Any]:
    rows = sorted(parsed["functions"],
                  key=lambda function_id: (parsed["functions"][function_id]["name"], function_id))
    result: dict[str, Any] = {}
    events = tuple(parsed["header"]["events"])
    for category, pattern in SELECTED_PATTERNS.items():
        matches = [function_id for function_id in rows
                   if pattern.search(parsed["functions"][function_id]["name"] or "")]
        detail = [selected_row(parsed, function_id) for function_id in matches]
        detail.sort(key=lambda item: (-item["self"]["Ir"], item["name"], item["id"]))
        result[category] = {
            "pattern": pattern.pattern,
            "function_count": len(detail),
            "rows": detail,
            "self_total": serializable_vector(
                vector_sum((parsed["functions"][item["id"]]["self"] for item in detail), events), events
            ),
            "inclusive_total": serializable_vector(
                vector_sum((function_inclusive(parsed, item["id"])
                            for item in detail), events), events
            ),
            "calls_in_total": sum(int(item["calls_in"]) for item in detail),
            "calls_out_total": sum(int(item["calls_out"]) for item in detail),
            "inclusive_rows_overlap": any(item["inclusive_overlaps"] for item in detail),
        }
    return result


def top_self_functions(parsed: Mapping[str, Any], limit: int = 24) -> list[dict[str, Any]]:
    events = tuple(parsed["header"]["events"])
    function_ids = sorted(
        parsed["functions"],
        key=lambda function_id: (-parsed["functions"][function_id]["self"][0],
                                 parsed["functions"][function_id]["name"], function_id),
    )
    result: list[dict[str, Any]] = []
    for function_id in function_ids[:limit]:
        view = function_view(parsed, function_id)
        result.append(
            {
                "id": function_id,
                "name": view["name"],
                "self": serializable_vector(view["self"], events),
                "direct_children_cost": serializable_vector(view["direct_children_cost"], events),
                "calls_in": int(view["calls_in"]),
                "calls_out": int(view["calls_out"]),
            }
        )
    return result


def conservation(parsed: Mapping[str, Any]) -> dict[str, Any]:
    events = tuple(parsed["header"]["events"])
    summary = parsed["header"]["summary"]
    totals = parsed["header"]["totals"]
    self_total = parsed["self_total"]
    matches_summary = vector_equal(self_total, summary)
    matches_totals = totals is not None and vector_equal(self_total, totals)
    require(matches_summary, f"{parsed['path']}: self totals do not match summary for every event")
    require(matches_totals, f"{parsed['path']}: self totals do not match totals for every event")
    return {
        "self_total": serializable_vector(self_total, events),
        "summary": serializable_vector(summary, events),
        "totals": serializable_vector(totals, events) if totals is not None else None,
        "self_matches_summary": matches_summary,
        "self_matches_totals": matches_totals,
        "all_counters_conserved": matches_summary and matches_totals,
    }


def raw_pair(row: Mapping[str, Any], plan: Mapping[str, Any]) -> dict[str, Any]:
    stem = str(row["stem"])
    artifacts = row["artifacts"]
    numbered_path = packet_path(artifacts[f"{stem}.callgrind.1"]["path"], f"{stem} positive dump")
    terminal_path = packet_path(artifacts[f"{stem}.callgrind"]["path"], f"{stem} termination dump")
    positive = parse_callgrind(numbered_path, expected_events=EVENTS)
    terminal = parse_callgrind(terminal_path, expected_events=EVENTS, allow_empty=True)
    require(positive["header"]["part"] == 1, f"{stem}: positive dump part changed")
    require(positive["header"]["trigger"] == "--dump-after=" + OWNER,
            f"{stem}: positive dump trigger changed")
    require(vector_nonzero(positive["header"]["summary"]), f"{stem}: positive summary is zero")
    require(terminal["header"]["part"] == 2,
            f"{stem}: termination dump part changed")
    require(terminal["header"]["trigger"] == "Program termination",
            f"{stem}: termination trigger changed")
    require(terminal["header"]["summary"] == vector_zero(EVENTS),
            f"{stem}: termination summary is not zero for every event")
    if terminal["header"]["totals"] is not None:
        require(terminal["header"]["totals"] == vector_zero(EVENTS),
                f"{stem}: termination totals are not zero for every event")
    unexpected = terminal_path.with_name(f"{stem}.callgrind.2")
    require(not unexpected.exists(), f"{stem}: unexpected additional numbered dump")
    positive_conservation = conservation(positive)
    terminal_conservation = {
        "self_total": serializable_vector(terminal["self_total"], EVENTS),
        "summary": serializable_vector(terminal["header"]["summary"], EVENTS),
        "totals": serializable_vector(terminal["header"]["totals"], EVENTS)
                   if terminal["header"]["totals"] is not None else None,
        "all_counters_conserved": terminal["self_total"] == vector_zero(EVENTS),
    }
    require(terminal_conservation["all_counters_conserved"],
            f"{stem}: termination self counters are not zero")
    owner = owner_attribution(positive)
    return {
        "file": relative(numbered_path),
        "sha256": positive["sha256"],
        "bytes": positive["bytes"],
        "termination": {
            "file": relative(terminal_path),
            "sha256": terminal["sha256"],
            "bytes": terminal["bytes"],
            "summary": serializable_vector(terminal["header"]["summary"], EVENTS),
        },
        "summary": serializable_vector(positive["header"]["summary"], EVENTS),
        "totals": serializable_vector(positive["header"]["totals"], EVENTS)
                   if positive["header"]["totals"] is not None else None,
        "owner": owner,
        "selected_symbols": selected_symbols(positive),
        "top_self_functions": top_self_functions(positive),
        "parser": positive["statistics"],
        "conservation": {
            "positive": positive_conservation,
            "termination": terminal_conservation,
        },
        "validation": {
            "events_exact": list(positive["header"]["events"]) == list(EVENTS),
            "positive_dump_nonzero": True,
            "termination_zero_all_events": True,
            "positive_self_totals_match_summary_and_totals": True,
            "termination_self_totals_zero": True,
            "owner_qualification_is_reported": True,
            "nested_costs_overlap": True,
            "fractions_not_reported": True,
        },
    }


def analyze() -> dict[str, Any]:
    plan, plan_file = validate_plan()
    rows, receipt_identity = receipt_rows(plan)
    profiles: list[dict[str, Any]] = []
    for row in rows:
        stem = row["stem"]
        analyzed = raw_pair(row, plan)
        analyzed.update(
            {
                "repeat": row["repeat"],
                "shape": row["shape"],
                "leg": row["leg"],
                "report": f"profiles/{stem}.json",
                "report_sha256": row["artifacts"][f"{stem}.json"]["sha256"],
                "binary": row["binary"],
                "receipt": {
                    "started": row["started"],
                    "ended": row["ended"],
                    "command": row["command"],
                },
            }
        )
        profiles.append(analyzed)
    owner_qualified = sum(1 for profile in profiles if profile["owner"]["owner_qualified"])
    owner_failures = [
        {
            "repeat": profile["repeat"],
            "shape": profile["shape"],
            "leg": profile["leg"],
            "failures": profile["owner"]["qualification_failures"],
        }
        for profile in profiles
        if not profile["owner"]["owner_qualified"]
    ]
    return {
        "schema": "litchi-0795-callgrind-analysis-v1",
        "packet": "change-0795",
        "plan": {
            "path": relative(plan_file),
            "sha256": c.sha(plan_file),
            "schema": plan["schema"],
            "owner": OWNER,
            "events": list(EVENTS),
            "cpu": plan["cpu"],
        },
        "receipts": receipt_identity,
        "profile_count": len(profiles),
        "profiles": profiles,
        "summary": {
            "profiles": len(profiles),
            "events": list(EVENTS),
            "all_counter_conservation_pass": all(
                profile["conservation"]["positive"]["all_counters_conserved"]
                and profile["conservation"]["termination"]["all_counters_conserved"]
                for profile in profiles
            ),
            "owner_qualified_profiles": owner_qualified,
            "owner_qualification_failures": owner_failures,
            "nested_costs_overlap": True,
            "fractions_reported": False,
        },
        "claims": [
            "Callgrind Ir, branch counts, and misprediction counts are guest-counter diagnostics, not native latency.",
            "All five counters are conserved from parsed self rows to each positive dump summary and totals line.",
            "Owner and descendant inclusive costs overlap; only the owner self plus immediate-child inclusive rows form a disjoint partition when owner qualification succeeds.",
            "Owner qualification failures are retained in the report and are not converted into a success claim.",
            "No phase fractions, native speedup, or production-adoption claim is made by this analysis.",
            "The allocation_named_functions group is a name-based function census; its incoming and outgoing graph calls are not allocator API call counts.",
        ],
    }


def markdown(report: Mapping[str, Any]) -> str:
    events = report["plan"]["events"]
    lines = [
        "# 0795 multi event Callgrind analysis",
        "",
        "This is an offline replay of the twelve retained Callgrind captures. "
        "The event columns are guest counter diagnostics and do not measure native latency.",
        "",
        "| Repeat | Shape | Leg | Ir | Bc | Bcm | Bi | Bim | Owner qualified |",
        "| ---: | --- | --- | ---: | ---: | ---: | ---: | ---: | --- |",
    ]
    for profile in report["profiles"]:
        summary = profile["summary"]
        owner = profile["owner"]
        lines.append(
            f"| {profile['repeat']} | {profile['shape']} | {profile['leg']} | "
            + " | ".join(f"{summary[event]:,}" for event in events)
            + f" | {'yes' if owner['owner_qualified'] else 'no'} |"
        )
    lines += [
        "",
        "For every positive dump the parser sums every function self vector and "
        "checks it against both `summary:` and `totals:` for all five events. "
        "The termination dump is required to be zero for every event.",
        "",
        "Immediate owner children are the only inclusive rows used in the "
        "disjoint partition. Descendant inclusive rows overlap and are retained "
        "for path inspection; no fractions are computed.",
        "",
        "The `allocation_named_functions` group is a name-based function census. "
        "Its incoming and outgoing graph calls are not allocator API call counts.",
        "",
        "## Owner qualification failures",
        "",
    ]
    failures = report["summary"]["owner_qualification_failures"]
    if failures:
        for failure in failures:
            lines.append(
                f"- `{failure['repeat']}-{failure['shape']}-{failure['leg']}`: "
                + "; ".join(failure["failures"])
            )
    else:
        lines.append("None.")
    lines += ["", "## Selected symbol accounting", ""]
    for profile in report["profiles"]:
        lines += [f"### {profile['repeat']}-{profile['shape']}-{profile['leg']}", ""]
        lines.append("| Group | Functions | Self Ir | Inclusive Ir (overlapping) | Calls in | Calls out |")
        lines.append("| --- | ---: | ---: | ---: | ---: | ---: |")
        for category, detail in profile["selected_symbols"].items():
            lines.append(
                f"| `{category}` | {detail['function_count']} | "
                f"{detail['self_total']['Ir']:,} | {detail['inclusive_total']['Ir']:,} | "
                f"{detail['calls_in_total']:,} | {detail['calls_out_total']:,} |"
            )
        lines.append("")
    lines += [
        "Raw dumps, termination dumps, reports, and receipt artifact hashes are "
        "bound in `callgrind-analysis.json`.",
        "",
    ]
    return "\n".join(lines)


def write_or_check(path: Path, content: str, check: bool) -> None:
    if check:
        require(path.is_file() and not path.is_symlink(), f"missing expected output: {relative(path)}")
        require(path.read_text(encoding="utf-8") == content,
                f"replayed output differs: {relative(path)}")
    else:
        path.write_text(content, encoding="utf-8")


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    mode = parser.add_mutually_exclusive_group()
    mode.add_argument("--write", action="store_true", help="write deterministic analysis outputs")
    mode.add_argument("--check", action="store_true", help="replay and compare existing outputs")
    args = parser.parse_args(argv)
    try:
        report = analyze()
        encoded = json.dumps(report, indent=2, sort_keys=True) + "\n"
        rendered = markdown(report)
        write_or_check(ANALYSIS_JSON, encoded, args.check)
        write_or_check(ANALYSIS_MD, rendered, args.check)
        print(f"0795 Callgrind analysis {'check' if args.check else 'write'} PASS", flush=True)
        return 0
    except (EvidenceError, CallgrindError, OSError, KeyError, TypeError, ValueError) as error:
        print(f"cg_analysis.py: error: {error}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
