#!/usr/bin/env python3
"""Reduce retained 0901 Callgrind dumps into XML-audit hot-loop evidence.

This helper is deliberately offline.  It imports the retained 0845 raw
Callgrind reader, validates the frozen 0903 custody bindings, and records raw
function/edge partitions.  It never starts a workload, compiler, profiler, or
external annotation tool.
"""

from __future__ import annotations

import argparse
import hashlib
import importlib.util
import json
import os
from pathlib import Path
import re
import sys
from typing import Any, Mapping


HERE = Path(__file__).resolve().parent
PLAN_PATH = HERE / "plan.json"
READER_PATH = HERE.parent / "change-0845" / "reader_v2.py"

OWNER = "litchi_xlsx::cell_values::source::SourceBackedEditor::publish_multi_commit_to_stream"
MEASURED_PARENT = "litchi_perf_baseline::run_xlsx_cell_values_edit_save"
VERIFY_REPLACEMENT = "xml_minifier::audit::verify_source_replacement"
VERIFY_POLICY = "xml_minifier::audit::verify_with_policy"
TARGETS = (
    VERIFY_REPLACEMENT,
    VERIFY_POLICY,
    "xml_minifier::audit::check_start",
    "xml_minifier::audit::inspect_attributes",
    "xml_minifier::audit::wellformed::scan_qname",
)
PROFILE_SUFFIX = re.compile(r"capture/(\d\d)-([a-z-]+)/callgrind\.out\.6$")
XML_SOURCE_MARKER = "/crates/xml-minifier/"


class EvidenceError(ValueError):
    """A missing, malformed, or contradictory retained artifact."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise EvidenceError(message)


def digest(path: Path) -> str:
    require(path.is_file() and not path.is_symlink(), f"missing regular file: {path}")
    value = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1 << 20), b""):
            value.update(chunk)
    return value.hexdigest()


def read_json(path: Path) -> Any:
    require(path.is_file() and not path.is_symlink(), f"missing JSON file: {path}")
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise EvidenceError(f"cannot read JSON {path}: {error}") from error


def descriptor(value: Any, label: str) -> dict[str, Any]:
    """Verify a frozen path/size/hash descriptor and return it unchanged."""

    require(isinstance(value, Mapping), f"{label} is not an object")
    path_value = value.get("path")
    require(isinstance(path_value, str) and path_value, f"{label}.path is missing")
    path = Path(path_value)
    require(path.is_file() and not path.is_symlink(), f"{label} file is missing: {path}")
    expected_bytes = value.get("bytes")
    expected_hash = value.get("sha256")
    require(type(expected_bytes) is int and expected_bytes >= 0,
            f"{label}.bytes is invalid")
    require(isinstance(expected_hash, str) and len(expected_hash) == 64,
            f"{label}.sha256 is invalid")
    actual_bytes = path.stat().st_size
    actual_hash = digest(path)
    require(actual_bytes == expected_bytes,
            f"{label}.bytes differs: {actual_bytes} != {expected_bytes}")
    require(actual_hash == expected_hash,
            f"{label}.sha256 differs: {actual_hash} != {expected_hash}")
    return dict(value)


def plan_descriptor(path: Path) -> dict[str, Any]:
    return {"path": str(path), "bytes": path.stat().st_size, "sha256": digest(path)}


def load_reader() -> Any:
    require(READER_PATH.is_file(), f"missing retained reader: {READER_PATH}")
    spec = importlib.util.spec_from_file_location("retained_callgrind_reader_0845", READER_PATH)
    require(spec is not None and spec.loader is not None,
            f"cannot load retained reader: {READER_PATH}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


def verify_prior_evidence(plan: Mapping[str, Any]) -> dict[str, Any]:
    source_descriptor = descriptor(plan.get("source_manifest"), "plan.source_manifest")
    source = read_json(Path(source_descriptor["path"]))
    require(isinstance(source, Mapping), "source manifest is not an object")
    require(source.get("baseline_commit") == "506aa2edae6666d9c76ae6827c2f5941a8e47896",
            "source manifest baseline commit differs")
    require(source.get("file_count") == 10771, "source manifest file count differs")
    files = source.get("files")
    require(isinstance(files, Mapping) and len(files) == 10771,
            "source manifest file inventory differs")

    prior_analysis_descriptor = descriptor(plan.get("prior_analysis"), "plan.prior_analysis")
    prior_analysis = read_json(Path(prior_analysis_descriptor["path"]))
    require(isinstance(prior_analysis, Mapping), "prior analysis is not an object")
    require(prior_analysis.get("schema") == "litchi.xlsx-publication-profile-analysis-0901.v1",
            "prior analysis schema differs")
    require(prior_analysis.get("status") == "pass" and prior_analysis.get("mode") == "full",
            "prior analysis status or mode differs")
    require(prior_analysis.get("owner") == OWNER
            and prior_analysis.get("measured_parent") == MEASURED_PARENT,
            "prior analysis owner binding differs")
    captures = prior_analysis.get("captures")
    require(isinstance(captures, list) and len(captures) == 4,
            "prior analysis capture count differs")

    prior_plan_descriptor = descriptor(plan.get("prior_plan"), "plan.prior_plan")
    prior_plan = read_json(Path(prior_plan_descriptor["path"]))
    require(isinstance(prior_plan, Mapping), "prior plan is not an object")
    require(prior_plan.get("schema") == "litchi.xlsx-publication-profile-0901.v1",
            "prior plan schema differs")
    require(prior_plan.get("source_files_verified") == 10771,
            "prior plan source inventory differs")
    require(prior_plan.get("owner") == OWNER
            and prior_plan.get("measured_parent") == MEASURED_PARENT,
            "prior plan owner binding differs")
    schedule = prior_plan.get("schedule")
    require(isinstance(schedule, list) and len(schedule) == 4,
            "prior plan schedule differs")
    expected_schedule = [
        {"index": 0, "repeat": 0, "shape": "medium"},
        {"index": 1, "repeat": 0, "shape": "dense-sparse"},
        {"index": 2, "repeat": 1, "shape": "dense-sparse"},
        {"index": 3, "repeat": 1, "shape": "medium"},
    ]
    require(schedule == expected_schedule, "prior plan schedule order differs")

    prior_review_descriptor = descriptor(plan.get("prior_review"), "plan.prior_review")
    prior_review = read_json(Path(prior_review_descriptor["path"]))
    require(isinstance(prior_review, Mapping), "prior review is not an object")
    canonical = prior_review.get("canonical_analysis")
    require(isinstance(canonical, Mapping), "prior review canonical analysis is missing")
    require(canonical.get("schema") == "litchi.xlsx-publication-profile-analysis-0901.v1"
            and canonical.get("status") == "pass"
            and canonical.get("captures") == 4,
            "prior review canonical binding differs")

    decision_descriptor = descriptor(plan.get("rejected_candidate_decision"),
                                     "plan.rejected_candidate_decision")
    decision = read_json(Path(decision_descriptor["path"]))
    require(isinstance(decision, Mapping), "rejected candidate decision is not an object")
    require(decision.get("production_retained") is False,
            "0902 rejected candidate was retained")
    require(isinstance(decision.get("performance_claim"), str)
            and "no adopted improvement" in decision["performance_claim"],
            "0902 decision does not record rejection")

    return {
        "source_manifest": source_descriptor,
        "source_inventory": {
            "baseline_commit": source["baseline_commit"],
            "file_count": source["file_count"],
            "files_entries": len(files),
            "inventory_matches_frozen_0894": True,
        },
        "prior_analysis": prior_analysis_descriptor,
        "prior_analysis_binding": {
            "schema": prior_analysis["schema"],
            "status": prior_analysis["status"],
            "mode": prior_analysis["mode"],
            "captures": len(captures),
        },
        "prior_plan": prior_plan_descriptor,
        "prior_review": prior_review_descriptor,
        "rejected_candidate_decision": decision_descriptor,
        "rejected_candidate_binding": {
            "production_retained": decision["production_retained"],
            "performance_claim": decision["performance_claim"],
        },
        "schedule": expected_schedule,
    }


def profile_edge(edge: Mapping[str, Any], names: Mapping[int, str]) -> dict[str, Any]:
    position = edge.get("position")
    if position == "*":
        position_kind = "wildcard"
    elif isinstance(position, str) and position.startswith(("+", "-")):
        position_kind = "relative"
    elif isinstance(position, str) and position.lstrip("+-").isdigit():
        position_kind = "absolute"
    else:
        position_kind = "unknown"
    return {
        "caller_id": edge["caller_id"],
        "caller": names.get(edge["caller_id"], "<unnamed>"),
        "callee_id": edge["callee_id"],
        "callee": names.get(edge["callee_id"], "<unnamed>"),
        "calls": edge["calls"],
        "inclusive_ir": edge["inclusive_ir"],
        "file": edge.get("file", ""),
        "position": position,
        "position_kind": position_kind,
        "reader_record_line": edge.get("line"),
    }


def all_incoming(parsed: Mapping[str, Any], target_id: int) -> list[dict[str, Any]]:
    names = parsed["names"]
    rows: list[dict[str, Any]] = []
    for function in parsed["functions"].values():
        for edge in function["edges"]:
            if edge["callee_id"] == target_id:
                rows.append(profile_edge(edge, names))
    return rows


def cost_rows(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Return edges carrying Ir, including a raw edge whose calls field is 0."""

    return [row for row in rows if row["inclusive_ir"] > 0]


def call_rows(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    return [row for row in rows if row["calls"] > 0]


def function_partition(parsed: Mapping[str, Any], function_id: int) -> dict[str, Any]:
    function = parsed["functions"][function_id]
    names = parsed["names"]
    edges = [profile_edge(edge, names) for edge in function["edges"]]
    incoming = all_incoming(parsed, function_id)
    direct_cost = cost_rows(edges)
    incoming_cost = cost_rows(incoming)
    direct_call_rows = call_rows(edges)
    incoming_call_rows = call_rows(incoming)
    direct_ir = sum(row["inclusive_ir"] for row in direct_cost)
    incoming_ir = sum(row["inclusive_ir"] for row in incoming_cost)
    direct_calls = sum(row["calls"] for row in direct_call_rows)
    incoming_calls = sum(row["calls"] for row in incoming_call_rows)
    self_ir = function["self_ir"]
    conservation = self_ir + direct_ir == incoming_ir
    require(conservation,
            f"function {function_id} {function.get('name', '')} does not conserve self plus direct edges")
    return {
        "id": function_id,
        "name": function.get("name") or names.get(function_id, "<unnamed>"),
        "file": function.get("file", ""),
        "records": function["records"],
        "self_ir": self_ir,
        "edges": edges,
        "edge_count": len(edges),
        "cost_edge_count": len(direct_cost),
        "positive_call_edge_count": len(direct_call_rows),
        "direct_children_ir": direct_ir,
        "direct_calls": direct_calls,
        "incoming_edges": incoming,
        "incoming_edge_count": len(incoming),
        "cost_incoming_edge_count": len(incoming_cost),
        "positive_call_incoming_edge_count": len(incoming_call_rows),
        "incoming_ir": incoming_ir,
        "incoming_calls": incoming_calls,
        "self_plus_direct_ir": self_ir + direct_ir,
        "conservation": {
            "self_plus_cost_edges_equals_cost_incoming": conservation,
            "zero_cost_direct_edges_excluded_from_cost_sum":
                len(edges) - len(direct_cost),
        },
    }


def exact_ids(parsed: Mapping[str, Any], name: str) -> list[int]:
    return sorted(function_id for function_id, function_name in parsed["names"].items()
                  if function_name == name and function_id in parsed["functions"])


def xml_source_file(path: str) -> bool:
    normalized = path.replace("\\", "/")
    return XML_SOURCE_MARKER in normalized


def xml_self_partition(parsed: Mapping[str, Any], owner_inclusive: int) -> dict[str, Any]:
    rows: list[dict[str, Any]] = []
    mismatches: list[dict[str, Any]] = []
    all_self = 0
    xml_self = 0
    for function_id, function in parsed["functions"].items():
        self_ir = function["self_ir"]
        all_self += self_ir
        name = function.get("name") or parsed["names"].get(function_id, "<unnamed>")
        file_name = function.get("file", "")
        by_name = name.startswith("xml_minifier::")
        by_file = xml_source_file(file_name)
        if by_name and not by_file:
            mismatches.append({"id": function_id, "name": name, "file": file_name,
                               "self_ir": self_ir})
        if by_file:
            xml_self += self_ir
            rows.append({"id": function_id, "name": name, "file": file_name,
                         "records": function["records"], "self_ir": self_ir})
    require(all_self == parsed["header"]["summary_ir"],
            f"{parsed['path']}: all function self costs do not conserve summary")
    require(not mismatches,
            f"{parsed['path']}: XML-named function lacks XML source attribution: {mismatches}")
    rows.sort(key=lambda row: (-row["self_ir"], row["id"]))
    remainder = owner_inclusive - xml_self
    require(remainder >= 0, f"{parsed['path']}: XML self cost exceeds publication owner")
    require(xml_self + remainder == owner_inclusive,
            f"{parsed['path']}: XML self plus remainder does not conserve owner")
    non_xml_self = all_self - xml_self
    require(non_xml_self == remainder,
            f"{parsed['path']}: non-XML self remainder differs from publication remainder")
    return {
        "functions": rows,
        "function_count": len(rows),
        "xml_self_ir": xml_self,
        "non_xml_self_remainder_ir": remainder,
        "all_exclusive_self_ir": all_self,
        "publication_owner_inclusive_ir": owner_inclusive,
        "validation": {
            "all_exclusive_self_conserves_summary": True,
            "xml_and_non_xml_self_are_disjoint": True,
            "xml_plus_remainder_conserves_publication_owner": True,
            "non_xml_self_equals_remainder": True,
        },
    }


def side_attribution(partitions: list[dict[str, Any]]) -> dict[str, Any]:
    evidence: list[dict[str, Any]] = []
    for partition in partitions:
        if partition["name"] != VERIFY_REPLACEMENT:
            continue
        for edge in partition["edges"]:
            if edge["callee"] == VERIFY_POLICY and edge["inclusive_ir"] > 0:
                evidence.append({
                    "function_id": partition["id"],
                    "callee_id": edge["callee_id"],
                    "calls": edge["calls"],
                    "inclusive_ir": edge["inclusive_ir"],
                    "file": edge["file"],
                    "position": edge["position"],
                    "position_kind": edge["position_kind"],
                    "reader_record_line": edge["reader_record_line"],
                })
    return {
        "status": "undetermined",
        "distinguishable": False,
        "method": "outgoing verify_source_replacement edge raw position/file",
        "evidence": evidence,
        "reason": (
            "The retained reader exposes the raw cost-position token and Callgrind file "
            "context, but no validated source-callsite decoder. Relative positions remain "
            "encoded tokens, absolute positions are not resolved here, and the measured "
            "function has no independently proven original-versus-replacement mapping."
        ),
        "no_side_assignment": True,
    }


def validate_profile(plan: Mapping[str, Any], descriptor_value: Mapping[str, Any],
                     expected: Mapping[str, Any], reader: Any) -> dict[str, Any]:
    frozen = descriptor(descriptor_value, f"plan.profiles[{expected['index']}]")
    path = Path(frozen["path"])
    match = PROFILE_SUFFIX.search(path.as_posix())
    require(match is not None, f"profile path does not match frozen capture layout: {path}")
    require(int(match.group(1)) == expected["index"] and match.group(2) == expected["shape"],
            f"profile schedule differs for {path}")
    parsed = reader.parse_raw(path)
    header = parsed["header"]
    require(header["part"] == 6, f"{path}: Callgrind part differs")
    require(header["trigger"] == "--dump-after=" + OWNER,
            f"{path}: Callgrind trigger differs")
    require(header["summary_ir"] == header["totals_ir"],
            f"{path}: summary and totals differ")
    command = header.get("command") or ""
    require("xlsx_source_backed_cell_values_one_percent_edit_save" in command,
            f"{path}: command case differs")
    require(("--xlsx-cell-crud-shape " + expected["shape"]) in command,
            f"{path}: command shape differs")

    owner_ids = exact_ids(parsed, OWNER)
    require(owner_ids, f"{path}: publication owner function is missing")
    owner_candidates: list[tuple[int, list[dict[str, Any]]]] = []
    for function_id in owner_ids:
        incoming = [row for row in all_incoming(parsed, function_id)
                    if row["calls"] > 0 and row["inclusive_ir"] > 0]
        selected = [row for row in incoming if row["caller"] == MEASURED_PARENT]
        if selected:
            owner_candidates.append((function_id, selected))
    require(len(owner_candidates) == 1,
            f"{path}: measured publication owner edge is ambiguous: {owner_candidates}")
    owner_id, selected_edges = owner_candidates[0]
    require(len(selected_edges) == 1 and selected_edges[0]["calls"] == 1,
            f"{path}: measured publication owner incoming edge is not one call")
    owner = function_partition(parsed, owner_id)
    require(owner["incoming_ir"] == header["summary_ir"],
            f"{path}: owner incoming Ir differs from Callgrind summary")
    require(owner["self_plus_direct_ir"] == selected_edges[0]["inclusive_ir"],
            f"{path}: owner self plus direct children differs from selected incoming edge")

    partitions: list[dict[str, Any]] = []
    for name in TARGETS:
        ids = exact_ids(parsed, name)
        require(ids, f"{path}: target function is missing: {name}")
        for function_id in ids:
            partition = function_partition(parsed, function_id)
            require(partition["name"] == name, f"{path}: target name changed for {function_id}")
            partitions.append(partition)

    xml_self = xml_self_partition(parsed, header["summary_ir"])
    return {
        "index": expected["index"],
        "repeat": expected["repeat"],
        "shape": expected["shape"],
        "profile": frozen,
        "header": header,
        "publication_owner": {
            "selected_id": owner_id,
            "all_matching_ids": owner_ids,
            "selected_incoming_edge": selected_edges[0],
            "partition": owner,
        },
        "targets": partitions,
        "audit_side_attribution": side_attribution(partitions),
        "xml_minifier_self_partition": xml_self,
        "validation": {
            "profile_hash_and_size": True,
            "one_measured_owner_edge": True,
            "owner_incoming_equals_summary": True,
            "owner_self_plus_direct_children_conserves": True,
            "target_function_partitions_conserve": True,
            "all_incoming_callers_retained": True,
            "nested_inclusive_costs_not_added_to_xml_self_sum": True,
        },
    }


def analyze() -> dict[str, Any]:
    plan = read_json(PLAN_PATH)
    require(isinstance(plan, Mapping), "0903 plan is not an object")
    require(plan.get("schema") == "litchi.xml-audit-hotloops-0903.v1",
            "0903 plan schema differs")
    require(plan.get("status") == "frozen", "0903 plan is not frozen")
    require(plan.get("source_files_verified") == 10771,
            "0903 plan source inventory differs")
    require(plan.get("worktrees_created") == 0, "0903 plan records a worktree")
    require(isinstance(plan.get("scope"), str) and "no new workload" in plan["scope"],
            "0903 scope is not offline")
    profiles = plan.get("profiles")
    require(isinstance(profiles, list) and len(profiles) == 4,
            "0903 profile list differs")
    reader_descriptor = descriptor(plan.get("reader"), "plan.reader")
    require(Path(reader_descriptor["path"]) == READER_PATH,
            "plan reader path differs")
    require(reader_descriptor["bytes"] == READER_PATH.stat().st_size
            and reader_descriptor["sha256"] == digest(READER_PATH),
            "plan reader descriptor differs")
    prior = verify_prior_evidence(plan)
    schedule = prior["schedule"]
    require([PROFILE_SUFFIX.search(str(item["path"])).group(1) for item in profiles]
            == [f"{row['index']:02d}" for row in schedule],
            "profile descriptor order differs")

    reader = load_reader()
    captures = [validate_profile(plan, value, schedule[index], reader)
                for index, value in enumerate(profiles)]
    return {
        "schema": "litchi.xml-audit-hotloops-analysis-0903.v1",
        "status": "pass",
        "plan": plan_descriptor(PLAN_PATH),
        "plan_binding": {
            "schema": plan["schema"],
            "status": plan["status"],
            "source_files_verified": plan["source_files_verified"],
            "worktrees_created": plan["worktrees_created"],
        },
        "retained_reader": reader_descriptor,
        "prior_evidence": prior,
        "owner": OWNER,
        "measured_parent": MEASURED_PARENT,
        "target_functions": list(TARGETS),
        "captures": captures,
        "scope": plan["scope"],
        "limitations": [
            "Callgrind Ir is guest-instruction attribution, not native latency, hardware cycles, allocation, RSS, or a speedup claim.",
            "Each target partition retains every raw direct and incoming edge, including zero-cost records; nested inclusive costs are never added across callpaths.",
            "XML-minifier totals use disjoint exclusive self costs classified by retained source file; the remainder is the non-XML exclusive self bucket.",
            "Relative edge positions remain raw Callgrind tokens and absolute positions are not decoded; no source-line attribution is claimed.",
            "Recorded Callgrind call counts are edge fields only; they are not XML token counts, branch coverage, or a basis for cost-per-call interpretation.",
            "The original-versus-replacement audit side cannot be distinguished from the retained outgoing edge position/file fields alone.",
            "Inherited 0901 brk warnings, demangled specialization aggregation, and source-debug attribution limits remain applicable.",
        ],
    }


def write_once(path: Path, value: Mapping[str, Any]) -> None:
    require(path.parent.is_dir(), f"output parent does not exist: {path.parent}")
    data = (json.dumps(value, indent=2, sort_keys=True) + "\n").encode("utf-8")
    try:
        fd = os.open(path, os.O_WRONLY | os.O_CREAT | os.O_EXCL, 0o644)
    except FileExistsError as error:
        raise EvidenceError(f"refusing to overwrite existing output: {path}") from error
    try:
        with os.fdopen(fd, "wb") as stream:
            stream.write(data)
    except Exception:
        try:
            path.unlink()
        except OSError:
            pass
        raise


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", type=Path, default=HERE / "analysis.json",
                        help="new analysis JSON path; existing files are never overwritten")
    args = parser.parse_args()
    try:
        result = analyze()
        write_once(args.output, result)
    except (EvidenceError, OSError, ImportError) as error:
        print(f"reduce.py: error: {error}", file=sys.stderr)
        return 2
    print(f"0903 XML-audit reduction passed: {args.output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
