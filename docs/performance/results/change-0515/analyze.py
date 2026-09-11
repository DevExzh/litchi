#!/usr/bin/env python3
"""Parse the four bounded 0515 Callgrind profiles.

The profile lane is diagnostic evidence only.  This parser deliberately uses
direct call edges from the selected ``Edit::commit`` body and its immediate
children.  It does not add inclusive rows, whole-process totals, or readback
work to the attribution.  Callgrind's ``--separate-callers=3`` spelling is
represented by an apostrophe in a function name; the base name and the full
context are retained separately so the check does not depend on numeric
context identifiers.
"""

from __future__ import annotations

import hashlib
import json
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any


HERE = Path(__file__).resolve().parent
PROFILE_LANES = ("commit-r1", "commit-r2", "compact-r1", "compact-r2")
COMMIT_LANES = ("commit-r1", "commit-r2")
COMPACTION_LANES = ("compact-r1", "compact-r2")

RUNNER = "litchi_perf_baseline::run_xlsx_update_commit"
COMMIT = "litchi_xlsx::workbook::edit::semantic::transaction::Edit::commit"
STORE = "litchi_xlsx::workbook::model::Worksheet::store"
PARSE = "litchi_xlsx::raw::worksheet::parse"
COMPACTION = "litchi_xlsx::raw::compact::changed_worksheet"
PARSER_PARSE = (
    "litchi_xlsx::raw::worksheet::codec::<impl "
    "litchi_xlsx::raw::worksheet::model::Parser>::parse"
)
READER_READ_EVENT = "quick_xml::reader::Reader<R>::read_event_impl"
READER_PROCESS_EVENT = "quick_xml::reader::ns_reader::NsReader<R>::process_event"
RESOLVER_SET_LEVEL = "quick_xml::name::NamespaceResolver::set_level"
RESOLVER_RESOLVE_EVENT = "quick_xml::name::NamespaceResolver::resolve_event"
PARSER_START = (
    "litchi_xlsx::raw::worksheet::codec::<impl "
    "litchi_xlsx::raw::worksheet::model::Parser>::start"
)
PARSER_MATERIALIZE = "litchi_xlsx::raw::worksheet::semantic::materialize"
EAGER_CHILDREN = (
    READER_READ_EVENT,
    READER_PROCESS_EVENT,
    RESOLVER_SET_LEVEL,
    RESOLVER_RESOLVE_EVENT,
    PARSER_START,
    PARSER_MATERIALIZE,
)
SHARED_READER_CHILDREN = (
    READER_READ_EVENT,
    READER_PROCESS_EVENT,
    RESOLVER_SET_LEVEL,
)
RETAINED_PARSER_CHILDREN = (
    RESOLVER_RESOLVE_EVENT,
    PARSER_START,
    PARSER_MATERIALIZE,
)
EAGER_CALLS = {
    READER_READ_EVENT: 1_969_194,
    READER_PROCESS_EVENT: 1_969_194,
    RESOLVER_SET_LEVEL: 787_986,
    RESOLVER_RESOLVE_EVENT: 1_969_194,
    PARSER_START: 787_980,
    PARSER_MATERIALIZE: 393_216,
}

SHA256_RE = re.compile(r"^[0-9a-f]{64}$")
SUMMARY_RE = re.compile(r"^summary:\s*(\d+)\s*$", re.MULTILINE)
FN_RE = re.compile(r"^fn=\((?P<id>\d+)\)(?:\s(?P<name>.*))?\s*$")
CFN_RE = re.compile(r"^cfn=\((?P<id>\d+)\)(?:\s(?P<name>.*))?\s*$")
CALLS_RE = re.compile(r"^calls=(?P<count>[\d,]+)(?:\s|$)")
EVENT_COST_RE = re.compile(r"^\s*(?:\*|[+-]?\d+)\s+(?P<cost>\d+)\s*$")
ANNOTATION_FUNCTION_RE = re.compile(
    r"^\s*[\d,]+\s+\([^)]*\)\s+\*\s+(?P<function>.+?)\s*$"
)
ANNOTATION_EDGE_RE = re.compile(
    r"^\s*(?P<cost>[\d,]+)\s+\(\s*(?P<percent>[-+]?\d+(?:\.\d+)?)%\)"
    r"\s+>\s+(?P<function>.+?)\s+\((?P<count>[\d,]+)x\)"
    r"(?:\s+\[[^]]*\])?\s*$"
)


class ProfileError(ValueError):
    """A malformed or incomplete profile artifact."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise ProfileError(message)


def sha(path: Path) -> str:
    digest = hashlib.sha256()
    try:
        with path.open("rb") as stream:
            for block in iter(lambda: stream.read(1024 * 1024), b""):
                digest.update(block)
    except OSError as error:
        raise ProfileError(f"cannot read {path}: {error}") from error
    return digest.hexdigest()


def text(path: Path) -> str:
    try:
        return path.read_text(encoding="utf-8", errors="strict")
    except (OSError, UnicodeError) as error:
        raise ProfileError(f"cannot read {path}: {error}") from error


def check_profile_files(lane: str) -> tuple[Path, Path, Path]:
    require(lane in PROFILE_LANES, f"unknown profile lane {lane!r}")
    raw = HERE / f"{lane}.out"
    inclusive = HERE / f"{lane}-inclusive.txt"
    exclusive = HERE / f"{lane}-exclusive.txt"
    for path in (raw, inclusive, exclusive):
        require(path.is_file() and not path.is_symlink(),
                f"{lane} profile artifact is missing or not regular: {path.name}")
        require(path.stat().st_size > 0, f"{lane} profile artifact is empty: {path.name}")
    return raw, inclusive, exclusive


def base_name(name: str) -> str:
    """Drop annotation prefixes and Callgrind caller-context suffixes."""

    value = name.strip()
    if value.startswith("???:"):
        value = value[4:]
    # callgrind_annotate appends the image path to some rows.  The path is
    # presentation metadata, not part of the demangled symbol.
    value = value.split(" [", 1)[0]
    # Shared-library rows can carry a source path before the symbol, for
    # example ``./string/...:__memcpy_avx_unaligned_erms``.
    if value.startswith("./") or value.startswith("/"):
        value = value.split(":", 1)[-1]
    # The apostrophe is Callgrind's context separator.  Rust demangled names
    # in this profile do not use it as part of a function name.
    return value.split("'", 1)[0]


def context_name(name: str) -> str:
    value = name.strip()
    if value.startswith("???:"):
        value = value[4:]
    value = value.split(" [", 1)[0]
    if value.startswith("./") or value.startswith("/"):
        value = value.split(":", 1)[-1]
    return value


@dataclass(frozen=True)
class RawEdge:
    parent: str
    child: str
    calls: int
    cost: int


@dataclass(frozen=True)
class AnnotationEdge:
    parent: str
    child: str
    calls: int
    cost: int
    percent: float


def parse_raw(path: Path) -> tuple[int, list[RawEdge]]:
    content = text(path)
    summaries = SUMMARY_RE.findall(content)
    require(len(summaries) == 1 and int(summaries[0]) > 0,
            f"{path.name} must contain one nonzero Callgrind summary")

    # Resolve all function IDs first.  cfn records are allowed to omit a
    # repeated name, which is common in Callgrind output.
    names: dict[str, str] = {}
    records: list[tuple[str, str, int, int]] = []
    current_fn: str | None = None
    current_cfn: str | None = None
    lines = content.splitlines()
    for line_number, line in enumerate(lines):
        match = FN_RE.match(line)
        if match:
            current_fn = match.group("id")
            current_cfn = None
            if match.group("name"):
                value = match.group("name").strip()
                previous = names.get(current_fn)
                require(previous is None or previous == value,
                        f"{path.name} changes function name for fn={current_fn}")
                names[current_fn] = value
            continue
        match = CFN_RE.match(line)
        if match:
            current_cfn = match.group("id")
            if match.group("name"):
                value = match.group("name").strip()
                previous = names.get(current_cfn)
                require(previous is None or previous == value,
                        f"{path.name} changes function name for cfn={current_cfn}")
                names[current_cfn] = value
            continue
        match = CALLS_RE.match(line)
        if match:
            require(current_fn is not None and current_cfn is not None,
                    f"{path.name} has a calls record without fn/cfn")
            cost: int | None = None
            for following in lines[line_number + 1:]:
                event = EVENT_COST_RE.match(following)
                if event:
                    cost = int(event.group("cost"))
                    break
                if not following.strip():
                    break
            require(cost is not None,
                    f"{path.name} has a calls record without an Ir cost")
            records.append((
                current_fn,
                current_cfn,
                int(match.group("count").replace(",", "")),
                cost,
            ))

    edges: list[RawEdge] = []
    for parent_id, child_id, calls, cost in records:
        # Callgrind can retain a zero-call record with non-zero collected cost
        # after a collection reset.  Keep those records so context membership
        # remains faithful; the proof helpers below require positive calls on
        # the selected direct edges and sum only those selected edges.
        require(calls >= 0,
                f"{path.name} has a negative call edge")
        require(parent_id in names and child_id in names,
                f"{path.name} has an unresolved fn/cfn mapping")
        edges.append(RawEdge(names[parent_id], names[child_id], calls, cost))
    require(edges, f"{path.name} has no call edges")
    return int(summaries[0]), edges


def parse_annotation(path: Path) -> list[AnnotationEdge]:
    lines = text(path).splitlines()
    edges: list[AnnotationEdge] = []
    parent: str | None = None
    for line in lines:
        function = ANNOTATION_FUNCTION_RE.match(line)
        if function:
            parent = function.group("function").strip()
            continue
        edge = ANNOTATION_EDGE_RE.match(line)
        if edge and parent is not None:
            edges.append(
                AnnotationEdge(
                    parent=parent,
                    child=edge.group("function").strip(),
                    calls=int(edge.group("count").replace(",", "")),
                    cost=int(edge.group("cost").replace(",", "")),
                    percent=float(edge.group("percent")),
                )
            )
    require(edges, f"{path.name} has no direct-callee annotation edges")
    return edges


def raw_matches(edges: list[RawEdge], parent: str, child: str) -> list[RawEdge]:
    return [edge for edge in edges
            if base_name(edge.parent) == parent and base_name(edge.child) == child]


def annotation_matches(edges: list[AnnotationEdge], parent: str, child: str) -> list[AnnotationEdge]:
    return [edge for edge in edges
            if base_name(edge.parent) == parent and base_name(edge.child) == child]


def direct_edge_summary(
    raw_edges: list[RawEdge],
    annotation_edges: list[AnnotationEdge],
    parent: str,
    child: str,
    expected_calls: int,
    label: str,
) -> dict[str, Any]:
    raw_records = raw_matches(raw_edges, parent, child)
    annotated_records = annotation_matches(annotation_edges, parent, child)
    # Retain zero-call records while parsing the profile, but only positive
    # call records establish a measured direct edge.
    raw = [edge for edge in raw_records if edge.calls > 0]
    annotated = [edge for edge in annotated_records if edge.calls > 0]
    require(raw, f"{label} is absent from raw Callgrind edges")
    require(annotated, f"{label} is absent from inclusive annotation")
    raw_total = sum(edge.calls for edge in raw)
    annotated_total = sum(edge.calls for edge in annotated)
    raw_cost = sum(edge.cost for edge in raw)
    annotated_cost = sum(edge.cost for edge in annotated)
    require(raw_total == expected_calls,
            f"{label} raw direct calls {raw_total} != {expected_calls}")
    require(annotated_total == expected_calls,
            f"{label} annotated direct calls {annotated_total} != {expected_calls}")
    require(raw_cost == annotated_cost,
            f"{label} raw Ir cost {raw_cost} != annotated Ir cost {annotated_cost}")
    require(all(edge.calls > 0 for edge in raw + annotated),
            f"{label} has a non-positive direct edge")
    return {
        "parent": parent,
        "child": child,
        "expected_calls": expected_calls,
        "raw_record_count": len(raw_records),
        "raw_edge_count": len(raw),
        "raw_calls": raw_total,
        "raw_ir": raw_cost,
        "raw_zero_call_ir": sum(edge.cost for edge in raw_records if edge.calls == 0),
        "annotated_record_count": len(annotated_records),
        "annotated_edge_count": len(annotated),
        "annotated_calls": annotated_total,
        "annotated_ir": annotated_cost,
        "annotated_edges": [
            {
                "parent_context": context_name(edge.parent),
                "child_context": context_name(edge.child),
                "calls": edge.calls,
                "inclusive_ir": edge.cost,
                "inclusive_percent": edge.percent,
            }
            for edge in annotated
        ],
    }


def changed_parser_parent(name: str) -> bool:
    """Match the changed-output parser context without a numeric context ID."""

    parts = context_name(name).split("'")
    return parts[:4] == [PARSER_PARSE, PARSE, COMMIT, RUNNER]


def direct_children_summary(
    raw_edges: list[RawEdge],
    annotation_edges: list[AnnotationEdge],
    parent_match: Any,
    children: tuple[str, ...] | None,
    label: str,
) -> dict[str, Any]:
    """Retain direct child costs from one bounded parent context.

    Raw Callgrind edges can split a child across call sites while
    callgrind_annotate presents one aggregate row.  Aggregate only within the
    same direct parent and compare both call counts and Ir costs with the
    annotation output.  Zero-call raw records remain visible in the retained
    record count but do not establish a measured child edge.
    """

    raw_records = [edge for edge in raw_edges if parent_match(edge.parent)]
    annotation_records = [edge for edge in annotation_edges
                          if parent_match(edge.parent)]
    require(raw_records, f"{label} has no raw direct children")
    require(annotation_records, f"{label} has no annotated direct children")
    raw_positive = [edge for edge in raw_records if edge.calls > 0]
    annotation_positive = [edge for edge in annotation_records if edge.calls > 0]
    if children is not None:
        expected = set(children)
        raw_positive = [edge for edge in raw_positive
                        if base_name(edge.child) in expected]
        annotation_positive = [edge for edge in annotation_positive
                               if base_name(edge.child) in expected]
    require(raw_positive, f"{label} has no positive-call raw children")
    require(annotation_positive, f"{label} has no positive-call annotated children")

    raw_names = {base_name(edge.child) for edge in raw_positive}
    annotation_names = {base_name(edge.child) for edge in annotation_positive}
    require(raw_names == annotation_names,
            f"{label} raw/annotation child sets differ")
    if children is not None:
        expected = set(children)
        require(raw_names == expected,
                f"{label} direct child set differs: {sorted(raw_names)!r}")

    result: dict[str, Any] = {}
    for child in sorted(raw_names):
        raw_selected = [edge for edge in raw_positive
                        if base_name(edge.child) == child]
        annotation_selected = [edge for edge in annotation_positive
                               if base_name(edge.child) == child]
        raw_calls = sum(edge.calls for edge in raw_selected)
        annotation_calls = sum(edge.calls for edge in annotation_selected)
        raw_ir = sum(edge.cost for edge in raw_selected)
        annotation_ir = sum(edge.cost for edge in annotation_selected)
        require(raw_calls == annotation_calls,
                f"{label} {child} raw/annotation call counts differ")
        require(raw_ir == annotation_ir,
                f"{label} {child} raw/annotation Ir costs differ")
        if child in EAGER_CALLS:
            require(raw_calls == EAGER_CALLS[child],
                    f"{label} {child} calls {raw_calls} != {EAGER_CALLS[child]}")
        result[child] = {
            "calls": raw_calls,
            "raw_ir": raw_ir,
            "annotated_ir": annotation_ir,
            "raw_record_count": len(raw_selected),
            "raw_zero_call_ir": sum(
                edge.cost for edge in raw_records
                if base_name(edge.child) == child and edge.calls == 0
            ),
            "annotated_record_count": len(annotation_selected),
            "parent_contexts": sorted({context_name(edge.parent)
                                        for edge in raw_selected}),
            "child_contexts": sorted({context_name(edge.child)
                                       for edge in annotation_selected}),
            "annotation_percent": sorted({edge.percent
                                           for edge in annotation_selected}),
        }
    return {
        "children": result,
        "direct_child_names": sorted(raw_names),
        "scope": "Positive direct child edges under the selected parent context; raw zero-call metadata retained but excluded from measured totals",
    }


def eager_parser_attribution(
    summary_ir: int,
    raw_edges: list[RawEdge],
    annotation_edges: list[AnnotationEdge],
) -> dict[str, Any]:
    selected = direct_children_summary(
        raw_edges,
        annotation_edges,
        changed_parser_parent,
        EAGER_CHILDREN,
        "changed-output eager parser",
    )
    children = selected["children"]
    shared_raw = sum(children[name]["raw_ir"] for name in SHARED_READER_CHILDREN)
    shared_annotated = sum(children[name]["annotated_ir"]
                           for name in SHARED_READER_CHILDREN)
    require(shared_raw == shared_annotated,
            "shared reader child Ir total differs between raw and annotation")
    retained = {name: children[name] for name in RETAINED_PARSER_CHILDREN}
    return {
        "direct_children": selected,
        "shared_reader_candidate_ceiling": {
            "children": list(SHARED_READER_CHILDREN),
            "raw_ir": shared_raw,
            "annotated_ir": shared_annotated,
            "percent_of_commit_profile": round(shared_raw * 100.0 / summary_ir, 2),
            "interpretation": "Upper-bound arithmetic over disjoint reader/resolver direct children; diagnostic only and not a speedup claim",
        },
        "retained_children": {
            "children": list(RETAINED_PARSER_CHILDREN),
            "details": retained,
            "scope": "Direct parser children retained in the profile; no removal attribution is inferred",
        },
    }


def compaction_children_attribution(
    raw_edges: list[RawEdge],
    annotation_edges: list[AnnotationEdge],
) -> dict[str, Any]:
    selected = direct_children_summary(
        raw_edges,
        annotation_edges,
        lambda name: base_name(name) == COMPACTION,
        None,
        "changed worksheet compaction",
    )
    return {
        "direct_children": selected,
        "scope": "All positive direct children of changed_worksheet in the compaction-only profile; no worksheet parse or readback work is included",
    }


def parse_context_attribution(
    raw_edges: list[RawEdge],
    annotation_edges: list[AnnotationEdge],
) -> dict[str, Any]:
    """Prove the two parser roles from direct edges under commit.

    A call to ``raw::worksheet::parse`` from ``Edit::commit`` is the
    changed-output parse.  A call from ``Worksheet::store`` is the source
    Store parse.  Restricting the child to this exact symbol and requiring the
    complete direct-parent set prevents inclusive root/readback work from
    entering either attribution.
    """

    raw_parse_records = [edge for edge in raw_edges
                         if base_name(edge.child) == PARSE]
    raw_parse_edges = [edge for edge in raw_parse_records if edge.calls > 0]
    require(raw_parse_edges, "raw Callgrind profile has no worksheet parse edges")
    raw_parents = {base_name(edge.parent) for edge in raw_parse_edges}
    require(raw_parents == {COMMIT, STORE},
            f"worksheet parse has unexpected direct parents: {sorted(raw_parents)!r}")
    annotated_parse_records = [edge for edge in annotation_edges
                               if base_name(edge.child) == PARSE]
    annotated_parse_edges = [edge for edge in annotated_parse_records
                             if edge.calls > 0]
    require(annotated_parse_edges, "inclusive annotation has no worksheet parse edges")
    annotation_parents = {base_name(edge.parent) for edge in annotated_parse_edges}
    require(annotation_parents == {COMMIT, STORE},
            f"annotated worksheet parse has unexpected direct parents: {sorted(annotation_parents)!r}")

    changed = direct_edge_summary(
        raw_edges, annotation_edges, COMMIT, PARSE, 6,
        "changed-output Edit::commit -> worksheet parse",
    )
    source = direct_edge_summary(
        raw_edges, annotation_edges, STORE, PARSE, 6,
        "source Worksheet::store -> worksheet parse",
    )
    # The two direct roles must remain separately visible in the profile.  We
    # retain the full context spellings rather than relying on context IDs.
    context_pairs = {
        (base_name(edge.parent), context_name(edge.child))
        for edge in annotated_parse_edges
    }
    require(len(context_pairs) >= 2,
            "worksheet parse direct edges do not retain two caller contexts")
    return {
        "changed_output_parse": changed,
        "source_store_parse": source,
        "direct_parent_set": sorted(raw_parents),
        "context_pairs": [
            {"parent": parent, "child_context": child_context}
            for parent, child_context in sorted(context_pairs)
        ],
        "scope": "Direct worksheet parse edges only: Edit::commit is changed output and Worksheet::store is source Store; no inclusive root or readback aggregation",
    }


def profile_result(lane: str) -> dict[str, Any]:
    raw_path, inclusive_path, exclusive_path = check_profile_files(lane)
    summary_ir, raw_edges = parse_raw(raw_path)
    inclusive_edges = parse_annotation(inclusive_path)
    # The exclusive annotation is retained and hash-bound by verify.py.  It
    # is read here to ensure it is valid UTF-8 and non-empty, but no exclusive
    # row is added to the inclusive attribution.
    _ = text(exclusive_path)

    runner_commit = direct_edge_summary(
        raw_edges, inclusive_edges, RUNNER, COMMIT, 3,
        "runner -> Edit::commit",
    )
    result: dict[str, Any] = {
        "summary_ir": summary_ir,
        "runner_to_commit": runner_commit,
        "raw_sha256": sha(raw_path),
        "inclusive_sha256": sha(inclusive_path),
        "exclusive_sha256": sha(exclusive_path),
    }
    if lane in COMMIT_LANES:
        result["parse_attribution"] = parse_context_attribution(raw_edges, inclusive_edges)
        result["eager_parser_attribution"] = eager_parser_attribution(
            summary_ir, raw_edges, inclusive_edges
        )
        result["mode"] = "commit-separate-callers-3"
    else:
        changed = direct_edge_summary(
            raw_edges, inclusive_edges, COMMIT, COMPACTION, 6,
            "Edit::commit -> changed worksheet compaction",
        )
        compact_edges = [edge for edge in raw_edges
                         if base_name(edge.child) == COMPACTION]
        require({base_name(edge.parent) for edge in compact_edges} == {COMMIT},
                "compaction has an unexpected direct parent")
        result["compaction"] = {
            "changed_worksheet": changed,
            "scope": "Direct changed_worksheet edges under Edit::commit only; six calls are two changed sheets across three commits",
        }
        result["compaction_children_attribution"] = compaction_children_attribution(
            raw_edges, inclusive_edges
        )
        result["mode"] = "compaction"
    return result


def analyze() -> dict[str, Any]:
    symbols_path = HERE / "profile-symbols.json"
    try:
        symbols = json.loads(symbols_path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise ProfileError(f"cannot load profile-symbols.json: {error}") from error
    require(isinstance(symbols, dict), "profile-symbols.json must be an object")
    require(symbols.get("commit") == f"*{COMMIT}*::commit" or
            symbols.get("commit") == "*litchi_xlsx::workbook::edit::semantic::transaction::Edit*::commit",
            "profile-symbols.json commit selector differs")
    require(symbols.get("runner") == f"*{RUNNER}",
            "profile-symbols.json runner selector differs")
    require(symbols.get("compaction") == f"*{COMPACTION}",
            "profile-symbols.json compaction selector differs")

    profiles = {lane: profile_result(lane) for lane in PROFILE_LANES}
    for first, second in (("commit-r1", "commit-r2"), ("compact-r1", "compact-r2")):
        left = profiles[first]
        right = profiles[second]
        require(left["runner_to_commit"]["raw_calls"] == right["runner_to_commit"]["raw_calls"] == 3,
                f"{first}/{second} commit call counts disagree")
        if first.startswith("commit"):
            for key in ("changed_output_parse", "source_store_parse"):
                require(left["parse_attribution"][key]["raw_calls"] ==
                        right["parse_attribution"][key]["raw_calls"] == 6,
                        f"{first}/{second} {key} counts disagree")
            require(left["eager_parser_attribution"]["direct_children"][
                        "direct_child_names"] == sorted(EAGER_CHILDREN),
                    f"{first} eager parser child order/set changed")
            require(right["eager_parser_attribution"]["direct_children"][
                        "direct_child_names"] == sorted(EAGER_CHILDREN),
                    f"{second} eager parser child order/set changed")
        else:
            require(left["compaction"]["changed_worksheet"]["raw_calls"] ==
                    right["compaction"]["changed_worksheet"]["raw_calls"] == 6,
                    f"{first}/{second} compaction counts disagree")
            require(left["compaction_children_attribution"]["direct_children"][
                        "direct_child_names"] == right[
                            "compaction_children_attribution"]["direct_children"][
                                "direct_child_names"],
                    f"{first}/{second} compaction child sets disagree")

    shared_r1 = profiles["commit-r1"]["eager_parser_attribution"][
        "shared_reader_candidate_ceiling"]["raw_ir"]
    shared_r2 = profiles["commit-r2"]["eager_parser_attribution"][
        "shared_reader_candidate_ceiling"]["raw_ir"]
    return {
        "schema": "litchi-0515-attribution-v1",
        "scope": "Diagnostic simulated-instruction attribution from four Callgrind profiles; direct selected edges only, no latency, RSS, allocation, hardware, or speedup claim",
        "profile_lanes": list(PROFILE_LANES),
        "profiles": profiles,
        "proof": {
            "exact_commit_calls_per_profile": 3,
            "exact_changed_output_parse_calls_per_commit_profile": 6,
            "exact_source_store_parse_calls_per_commit_profile": 6,
            "exact_changed_worksheet_calls_per_compaction_profile": 6,
            "context_separated": True,
            "inclusive_rows_not_summed": True,
            "readback_and_root_work_excluded": True,
            "shared_reader_candidate_ceiling_ir_by_repeat": {
                "commit-r1": shared_r1,
                "commit-r2": shared_r2,
            },
            "shared_reader_candidate_ceiling_repeat_delta_ir": shared_r2 - shared_r1,
            "shared_reader_candidate_ceiling_repeat_delta_scope": "Same-build Callgrind repeat drift; per-profile raw and annotation costs still agree exactly",
        },
        "performance_claim": "none",
        "claim_authorized": False,
    }


def main() -> int:
    try:
        print(json.dumps(analyze(), indent=2, sort_keys=True))
    except ProfileError as error:
        print(f"profile analysis failed: {error}")
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
