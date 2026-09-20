#!/usr/bin/env python3
"""Validate and attribute the 0715 DOCX counting-publication profiles.

The profiler emits one Callgrind file for each dump part and a final file at
the stem path.  Setup publications are deliberately retained.  The measured
part is selected only from the positive raw incoming edge whose caller is
``ordinary_save::run_case``; file numbering and timing order do not select a
part.  ``summarize_profile`` is intentionally independent of the capture
receipt and can therefore be used by a packet-level analyzer or directly from
the command line.

The only disjoint attribution is the selected owner's self Ir plus its
immediate child edge Ir.  Recursive OPC and write/ZIP rows are retained as
overlapping diagnostics so that a caller can rank materialization and
compression without adding nested rows to the owner total.
"""

from __future__ import annotations

import argparse
from collections import defaultdict
import hashlib
import json
from pathlib import Path
import re
import sys
from typing import Any, Iterable, NoReturn, Sequence


MEASURED_PARENT = "litchi_perf_baseline::ordinary_save::run_case"
DEFAULT_OWNER = (
    "litchi_docx::package::codec::<impl "
    "litchi_docx::package::model::Package>::write_plain"
)
ATOMIC_SETUP_PARENT = "litchi_opc::atomic::replace_with_impl"
GENERATED_SETUP_PARENTS = {
    "litchi_perf_baseline::semantic_docx_bytes": 1,
    ATOMIC_SETUP_PARENT: 4,
}
REAL_SETUP_PARENTS = {ATOMIC_SETUP_PARENT: 4}

FUNCTION_RE = re.compile(r"^(fn|cfn)=\((\d+)\)(?:\s+(.*))?$")
CALLS_RE = re.compile(r"^calls=([0-9][0-9,]*)")
PART_RE = re.compile(r"^part:\s*(\d+)\s*$")
TRIGGER_RE = re.compile(r"^desc:\s*Trigger:\s*(.*?)\s*$")
EVENTS_RE = re.compile(r"^events:\s*(.*?)\s*$")
SUMMARY_RE = re.compile(r"^summary:\s*([0-9][0-9,]*)\s*$")
INTEGER_RE = re.compile(r"^[0-9][0-9,]*$")
LOCATION_RE = re.compile(r"^(?:\*|[+-]?[0-9]+)$")
NUMBERED_SUFFIX_RE = re.compile(r"\.callgrind\.(\d+)$")
STEM_SUFFIX_RE = re.compile(r"\.callgrind$")

INTERESTING_TOKENS = (
    "litchi_opc::",
    "packagewriter",
    "opcpackage",
    "write_plain",
    "write_counted",
    "write_to",
    "to_stream",
    "serialize",
    "material",
    "zip",
    "deflate",
    "zlib",
    "compress",
    "preserv",
)


class ProfileError(ValueError):
    """A missing, malformed, or contradictory Callgrind profile."""


def fail(message: str) -> NoReturn:
    raise ProfileError(message)


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    try:
        with path.open("rb") as stream:
            for block in iter(lambda: stream.read(1024 * 1024), b""):
                digest.update(block)
    except OSError as error:
        fail(f"cannot hash {path}: {error}")
    return digest.hexdigest()


def read_text(path: Path) -> str:
    try:
        return path.read_text(encoding="utf-8", errors="replace")
    except OSError as error:
        fail(f"cannot read {path}: {error}")


def _part_suffix(path: Path) -> int | None:
    match = NUMBERED_SUFFIX_RE.search(path.name)
    return int(match.group(1)) if match else None


def _is_stem(path: Path) -> bool:
    return STEM_SUFFIX_RE.search(path.name) is not None


def _profile_paths(paths: Iterable[str | Path]) -> tuple[list[Path], Path]:
    """Normalize an explicit profile file list and identify its terminal file."""

    values = [Path(value) for value in paths]
    require(values, "profile path list is empty")
    unique: dict[str, Path] = {}
    for path in values:
        key = str(path.resolve())
        require(path.is_file() and not path.is_symlink(), f"profile file is missing: {path}")
        unique[key] = path
    values = list(unique.values())
    numbered = [path for path in values if _part_suffix(path) is not None]
    terminal = [path for path in values if _is_stem(path)]
    require(len(terminal) == 1, "profile path list must contain one .callgrind terminal file")
    require(numbered, "profile path list contains no numbered Callgrind parts")
    numbered.sort(key=lambda path: _part_suffix(path) or 0)
    parts = [_part_suffix(path) for path in numbered]
    require(parts == list(range(1, len(parts) + 1)),
            f"numbered Callgrind parts are not contiguous: {parts}")
    stem = terminal[0]
    require(all(path.parent == stem.parent and path.name.startswith(stem.name + ".")
                for path in numbered),
            "numbered Callgrind parts do not share the terminal stem")
    return numbered, stem


def paths_for_stem(stem: str | Path) -> list[Path]:
    """Return ``stem`` and all contiguous numbered parts for a capture stem."""

    path = Path(stem)
    if _part_suffix(path) is not None:
        name = path.name.rsplit(".", 1)[0]
        path = path.with_name(name)
    require(_is_stem(path), f"not a .callgrind stem: {path}")
    candidates = [path]
    candidates.extend(sorted(
        (item for item in path.parent.glob(path.name + ".*")
         if item.is_file() and not item.is_symlink() and _part_suffix(item) is not None),
        key=lambda item: _part_suffix(item) or 0,
    ))
    return candidates


def _event_cost(stripped: str, event_index: int) -> int | None:
    """Read one raw Callgrind cost line for the selected event."""

    fields = stripped.split()
    if len(fields) < 2 or not LOCATION_RE.fullmatch(fields[0]):
        return None
    costs = fields[1:]
    if event_index >= len(costs) or not INTEGER_RE.fullmatch(costs[event_index]):
        return None
    return int(costs[event_index].replace(",", ""))


def _function_names(lines: Sequence[str]) -> dict[int, str]:
    names: dict[int, str] = {}
    for line in lines:
        match = FUNCTION_RE.match(line.strip())
        if not match or not match.group(3):
            continue
        identifier = int(match.group(2))
        name = match.group(3)
        previous = names.get(identifier)
        require(previous is None or previous == name,
                f"function id {identifier} has conflicting names: {previous!r}, {name!r}")
        names[identifier] = name
    return names


def _parse_raw(path: Path) -> dict[str, Any]:
    """Parse headers, function self costs, and raw caller-to-callee edges."""

    text = read_text(path)
    lines = text.splitlines()
    names = _function_names(lines)
    events: list[str] = []
    parts: list[int] = []
    triggers: list[str] = []
    summaries: list[int] = []
    for line in lines:
        stripped = line.strip()
        if match := EVENTS_RE.match(stripped):
            events.append(match.group(1))
        if match := PART_RE.match(stripped):
            parts.append(int(match.group(1)))
        if match := TRIGGER_RE.match(stripped):
            triggers.append(match.group(1))
        if match := SUMMARY_RE.match(stripped):
            summaries.append(int(match.group(1).replace(",", "")))
    require(events and events[-1].split() == ["Ir"],
            f"{path}: expected exactly the Ir event")
    require(len(parts) == 1, f"{path}: expected one part header, got {parts}")
    require(len(triggers) == 1, f"{path}: expected one Trigger header, got {triggers}")
    require(len(summaries) == 1, f"{path}: expected one summary header, got {summaries}")
    event_index = events[-1].split().index("Ir")

    functions: dict[str, dict[str, Any]] = {}
    edges: list[dict[str, Any]] = []
    current_name: str | None = None
    pending_id: int | None = None
    pending_calls: int | None = None
    pending_ir = 0

    def function(name: str) -> dict[str, Any]:
        return functions.setdefault(name, {"name": name, "self_ir": 0, "edges": []})

    def flush_child() -> None:
        nonlocal pending_id, pending_calls, pending_ir
        if pending_id is not None and pending_calls is not None:
            callee = names.get(pending_id, "")
            require(callee, f"{path}: cfn id {pending_id} has no name")
            if pending_calls > 0 and pending_ir > 0 and current_name:
                edge = {
                    "caller": current_name,
                    "callee": callee,
                    "calls": pending_calls,
                    "inclusive_ir": pending_ir,
                }
                edges.append(edge)
                function(current_name)["edges"].append(edge)
        pending_id = None
        pending_calls = None
        pending_ir = 0

    for line in lines:
        stripped = line.strip()
        match = FUNCTION_RE.match(stripped)
        if match:
            if match.group(1) == "fn":
                flush_child()
                identifier = int(match.group(2))
                current_name = names.get(identifier)
                if current_name:
                    function(current_name)
            else:
                flush_child()
                pending_id = int(match.group(2))
                pending_calls = None
            continue
        if call_match := CALLS_RE.match(stripped):
            require(pending_id is not None, f"{path}: calls record outside cfn")
            pending_calls = int(call_match.group(1).replace(",", ""))
            continue
        if current_name is None:
            continue
        cost = _event_cost(stripped, event_index)
        if cost is None:
            continue
        if pending_id is not None:
            pending_ir += cost
        else:
            function(current_name)["self_ir"] += cost
    flush_child()

    for value in functions.values():
        value["edges"] = list(value["edges"])
    return {
        "path": path,
        "part": parts[0],
        "trigger": triggers[0],
        "summary_ir": summaries[0],
        "events": ["Ir"],
        "functions": functions,
        "edges": edges,
    }


def _edge_summary(edges: Sequence[dict[str, Any]], target: str) -> dict[str, Any]:
    selected = [edge for edge in edges if edge["callee"] == target and edge["calls"] > 0
                and edge["inclusive_ir"] > 0]
    by_caller: dict[str, dict[str, Any]] = {}
    for edge in selected:
        caller = edge["caller"]
        row = by_caller.setdefault(caller, {
            "caller": caller, "calls": 0, "inclusive_ir": 0, "edge_count": 0,
        })
        row["calls"] += edge["calls"]
        row["inclusive_ir"] += edge["inclusive_ir"]
        row["edge_count"] += 1
    return {
        "target": target,
        "positive_edge_count": len(selected),
        "calls": sum(edge["calls"] for edge in selected),
        "inclusive_ir": sum(edge["inclusive_ir"] for edge in selected),
        "edges": [dict(edge) for edge in selected],
        "by_caller": [by_caller[key] for key in sorted(by_caller)],
    }


def _direct_children(parsed: dict[str, Any], owner: str) -> tuple[int, list[dict[str, Any]]]:
    owner_functions = parsed["functions"].get(owner)
    require(owner_functions is not None, f"{parsed['path']}: owner function block is missing")
    by_name: dict[str, dict[str, Any]] = {}
    for edge in owner_functions["edges"]:
        row = by_name.setdefault(edge["callee"], {
            "name": edge["callee"], "calls": 0, "inclusive_ir": 0, "edge_count": 0,
        })
        row["calls"] += edge["calls"]
        row["inclusive_ir"] += edge["inclusive_ir"]
        row["edge_count"] += 1
    direct = sorted(by_name.values(), key=lambda row: (-row["inclusive_ir"], row["name"]))
    return owner_functions["self_ir"], direct


def _nested_serialization(parsed: dict[str, Any], owner: str, depth_limit: int = 5) -> dict[str, Any]:
    """Return overlapping graph rows useful for materialization/ZIP ranking."""

    by_caller: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for edge in parsed["edges"]:
        by_caller[edge["caller"]].append(edge)
    for values in by_caller.values():
        values.sort(key=lambda row: (-row["inclusive_ir"], row["callee"]))

    rows: list[dict[str, Any]] = []
    stack: list[tuple[str, int, tuple[str, ...]]] = [(owner, 0, (owner,))]
    while stack:
        caller, depth, ancestry = stack.pop()
        if depth >= depth_limit:
            continue
        for edge in by_caller.get(caller, []):
            callee = edge["callee"]
            row = {
                "depth": depth + 1,
                "caller": caller,
                "callee": callee,
                "calls": edge["calls"],
                "inclusive_ir": edge["inclusive_ir"],
            }
            lowered = f"{caller}\n{callee}".lower()
            if any(token in lowered for token in INTERESTING_TOKENS):
                rows.append(row)
            if callee not in ancestry:
                stack.append((callee, depth + 1, ancestry + (callee,)))
            # A malformed or unusually expansive profile must not turn the
            # evidence analyzer into an unbounded graph walk.
            require(len(rows) <= 4096, f"{parsed['path']}: nested graph exceeds 4096 rows")

    by_name: dict[str, dict[str, Any]] = {}
    for row in rows:
        value = by_name.setdefault(row["callee"], {
            "name": row["callee"], "min_depth": row["depth"],
            "max_depth": row["depth"], "edge_count": 0,
            "inclusive_ir_observations": [],
        })
        value["min_depth"] = min(value["min_depth"], row["depth"])
        value["max_depth"] = max(value["max_depth"], row["depth"])
        value["edge_count"] += 1
        value["inclusive_ir_observations"].append(row["inclusive_ir"])
    aggregate = sorted(by_name.values(), key=lambda row: (-max(row["inclusive_ir_observations"]), row["name"]))
    return {
        "depth_limit": depth_limit,
        "overlapping": True,
        "rows": rows,
        "by_callee": aggregate,
        "ranking_note": "Nested inclusive Ir observations overlap and must not be summed with the owner or direct-child partition.",
    }


def _part_row(parsed: dict[str, Any], owner: str, measured: bool) -> dict[str, Any]:
    incoming = _edge_summary(parsed["edges"], owner)
    require(incoming["positive_edge_count"] > 0,
            f"{parsed['path']}: owner has no positive incoming raw edge")
    require(incoming["inclusive_ir"] == parsed["summary_ir"],
            f"{parsed['path']}: owner incoming Ir {incoming['inclusive_ir']} != summary {parsed['summary_ir']}")
    measured_edges = [edge for edge in incoming["edges"] if edge["caller"] == MEASURED_PARENT]
    if measured:
        require(len(measured_edges) == 1 and measured_edges[0]["calls"] == 1,
                f"{parsed['path']}: measured owner edge is not one positive run_case call")
    else:
        require(not measured_edges,
                f"{parsed['path']}: setup part contains a run_case owner edge")
    return {
        "file": str(parsed["path"]),
        "part": parsed["part"],
        "sha256": sha256(parsed["path"]),
        "summary_ir": parsed["summary_ir"],
        "trigger": parsed["trigger"],
        "role": "measured" if measured else "setup",
        "owner_incoming": incoming,
        "selection_edge": measured_edges[0] if measured_edges else None,
        "function_count": len(parsed["functions"]),
    }


def summarize_profile(
    paths: Iterable[str | Path], expected_owner: str = DEFAULT_OWNER,
    expected_generated: bool = True,
) -> dict[str, Any]:
    """Validate one counting-publication profile and return attribution data.

    ``paths`` must contain the terminal ``*.callgrind`` file and every
    contiguous numbered part ``*.callgrind.1`` through the last part.  The
    generated corpus has five setup parts plus one measured part; the real
    fixture has four setup parts plus one measured part.  The return value is
    JSON-serializable and keeps all setup/terminal evidence.
    """

    require(isinstance(expected_owner, str) and expected_owner,
            "expected owner is empty")
    if isinstance(paths, (str, Path)):
        paths = paths_for_stem(paths)
    numbered_paths, terminal_path = _profile_paths(paths)
    expected_count = 6 if expected_generated else 5
    require(len(numbered_paths) == expected_count,
            f"expected {expected_count} numbered parts, got {len(numbered_paths)}")
    parsed_parts = [_parse_raw(path) for path in numbered_paths]
    owner_triggers = f"--dump-after={expected_owner}"
    for index, parsed in enumerate(parsed_parts, start=1):
        require(parsed["part"] == index,
                f"{parsed['path']}: header part {parsed['part']} != suffix {index}")
        require(parsed["trigger"] == owner_triggers,
                f"{parsed['path']}: Trigger does not identify expected owner")
    measured_indices = []
    incoming_rows = []
    for index, parsed in enumerate(parsed_parts):
        incoming = _edge_summary(parsed["edges"], expected_owner)
        matching = [edge for edge in incoming["edges"] if edge["caller"] == MEASURED_PARENT]
        if matching:
            measured_indices.append(index)
        incoming_rows.append(matching)
    require(len(measured_indices) == 1,
            f"expected one measured run_case owner edge, got {len(measured_indices)}")
    measured_index = measured_indices[0]
    rows = [_part_row(parsed, expected_owner, index == measured_index)
            for index, parsed in enumerate(parsed_parts)]

    expected_setup = GENERATED_SETUP_PARENTS if expected_generated else REAL_SETUP_PARENTS
    observed_setup: dict[str, int] = defaultdict(int)
    for row in rows:
        if row["role"] != "setup":
            continue
        for edge in row["owner_incoming"]["edges"]:
            observed_setup[edge["caller"]] += 1
            require(edge["caller"] in expected_setup,
                    f"{row['file']}: unexpected setup owner caller {edge['caller']!r}")
    require(dict(observed_setup) == expected_setup,
            f"setup caller counts {dict(observed_setup)} != {expected_setup}")

    terminal = _parse_raw(terminal_path)
    require(terminal["part"] == expected_count + 1,
            f"{terminal_path}: terminal part {terminal['part']} does not follow numbered parts")
    require(terminal["trigger"] == "Program termination",
            f"{terminal_path}: terminal Trigger is not Program termination")
    require(terminal["summary_ir"] == 0,
            f"{terminal_path}: terminal Ir is {terminal['summary_ir']}, expected zero")
    terminal_owner = _edge_summary(terminal["edges"], expected_owner)
    require(terminal_owner["positive_edge_count"] == 0,
            f"{terminal_path}: terminal contains a positive owner edge")

    selected_parsed = parsed_parts[measured_index]
    owner_self, direct_children = _direct_children(selected_parsed, expected_owner)
    owner_ir = selected_parsed["summary_ir"]
    direct_ir = sum(row["inclusive_ir"] for row in direct_children)
    require(owner_self + direct_ir == owner_ir,
            f"{selected_parsed['path']}: self {owner_self} + direct {direct_ir} != owner {owner_ir}")
    nested = _nested_serialization(selected_parsed, expected_owner)

    return {
        "schema_version": 1,
        "status": "pass",
        "owner": expected_owner,
        "measured_parent": MEASURED_PARENT,
        "expected_generated": expected_generated,
        "numbered_part_count": expected_count,
        "setup_part_count": expected_count - 1,
        "parts": rows,
        "terminal": {
            "file": str(terminal_path),
            "part": terminal["part"],
            "sha256": sha256(terminal_path),
            "summary_ir": terminal["summary_ir"],
            "trigger": terminal["trigger"],
            "validation": {"zero_ir": True, "follows_numbered_parts": True},
        },
        "selected": {
            "file": str(selected_parsed["path"]),
            "part": selected_parsed["part"],
            "selected_by": "sole positive raw incoming edge from ordinary_save::run_case",
            "incoming_edge": rows[measured_index]["selection_edge"],
        },
        "owner_ir": owner_ir,
        "self_ir": owner_self,
        "direct_children": direct_children,
        "direct_child_partition": {
            "self_ir": owner_self,
            "direct_children_ir": direct_ir,
            "owner_ir": owner_ir,
            "disjoint": True,
            "equation": "self_ir + sum(immediate direct-child inclusive Ir) = owner inclusive Ir",
        },
        "nested_opc_write_serialization": nested,
        "setup_callers": dict(sorted(observed_setup.items())),
        "validation": {
            "all_numbered_parts_retained": True,
            "measured_selected_by_raw_run_case_edge": True,
            "setup_counts_match_corpus": True,
            "terminal_zero_ir": True,
            "direct_children_reconstruct_owner": True,
            "nested_rows_are_overlapping_diagnostics": True,
        },
        "limitations": [
            "Callgrind Ir is guest-instruction attribution, not native latency, hardware counters, allocation counts, or RSS.",
            "Nested OPC/write/ZIP rows are inclusive and overlapping; only self plus immediate children is a disjoint partition.",
        ],
    }


def analyze_profile(stem: str | Path, corpus_id: str,
                    plan_profile: dict[str, Any]) -> dict[str, Any]:
    """Analyze one frozen 0715 stem while binding it to ``plan.profile``.

    This small adapter is the intended packet-level API.  It leaves source,
    binary, receipt, and native-parity custody to the caller, while making the
    owner, setup caller counts, and numbered-part policy from ``plan.json``
    part of the profile result.
    """

    require(isinstance(plan_profile, dict), "plan.profile is not an object")
    owner = plan_profile.get("owner")
    require(isinstance(owner, str) and owner, "plan.profile owner is missing")
    require(plan_profile.get("measured_parent") == MEASURED_PARENT,
            "plan.profile measured parent differs")
    setup_parents = plan_profile.get("setup_parents")
    require(isinstance(setup_parents, dict), "plan.profile setup parents are missing")
    require(corpus_id in ("generated", "numbered-list"),
            f"unsupported profile corpus: {corpus_id}")
    expected_generated = corpus_id == "generated"
    expected_setup = GENERATED_SETUP_PARENTS if expected_generated else REAL_SETUP_PARENTS
    require(setup_parents.get(corpus_id) == expected_setup,
            f"plan.profile setup policy differs for {corpus_id}")
    numbered_parts = plan_profile.get("numbered_parts")
    require(isinstance(numbered_parts, dict)
            and numbered_parts.get(corpus_id) == (6 if expected_generated else 5),
            f"plan.profile numbered-part policy differs for {corpus_id}")
    result = summarize_profile(paths_for_stem(stem), owner, expected_generated)
    result["corpus_id"] = corpus_id
    result["plan_profile_binding"] = {
        "owner": owner,
        "measured_parent": MEASURED_PARENT,
        "setup_parents": expected_setup,
        "numbered_parts": numbered_parts[corpus_id],
    }
    return result


def _cli_paths(values: list[Path]) -> list[Path]:
    require(values, "provide a .callgrind stem or profile files")
    if len(values) == 1:
        path = values[0]
        if _is_stem(path) or _part_suffix(path) is not None:
            return paths_for_stem(path)
    return values


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("paths", nargs="+", type=Path,
                        help="a .callgrind stem or the terminal plus numbered parts")
    parser.add_argument("--owner", default=DEFAULT_OWNER)
    corpus = parser.add_mutually_exclusive_group()
    corpus.add_argument("--generated", action="store_true",
                        help="expect five setup parts and one measured part")
    corpus.add_argument("--real", action="store_true",
                        help="expect four setup parts and one measured part")
    parser.add_argument("--output", type=Path,
                        help="write JSON instead of printing it")
    args = parser.parse_args(argv)
    try:
        paths = _cli_paths(args.paths)
        expected_generated = not args.real
        value = summarize_profile(paths, args.owner, expected_generated)
        encoded = json.dumps(value, indent=2, sort_keys=True) + "\n"
        if args.output:
            require(not args.output.exists(), f"output already exists: {args.output}")
            args.output.write_text(encoded, encoding="utf-8")
        else:
            print(encoded, end="")
        return 0
    except (ProfileError, OSError, ValueError) as error:
        print(f"analyze_profiles.py: error: {error}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
