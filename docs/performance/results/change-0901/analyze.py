#!/usr/bin/env python3
"""Fail-closed offline analysis of the 0901 Callgrind publication captures.

The retained 0845 reader owns the Callgrind grammar.  This packet-specific
adapter adds the evidence rules needed by 0901: the publication method has
several monomorphized symbols with the same demangled name, so an owner is
selected by its positive incoming edge from the measured runner.  Lifecycle
calls are retained and classified separately.  Only immediate children of a
selected owner are eligible for a disjoint partition; nested XML, writer, and
compression rows are reported as overlapping diagnostics.

This module never runs the native workload or a profiler.  It may run
``callgrind_annotate`` over already-retained dumps when ``--annotate`` is
requested (the default for a full analysis); the raw parser and first-child
admission are otherwise entirely offline.
"""

from __future__ import annotations

import argparse
import hashlib
import importlib.util
import json
import os
from pathlib import Path
import re
import subprocess
import sys
from typing import Any, Mapping, Sequence


HERE = Path(__file__).resolve().parent
READER_PATH = HERE.parent / "change-0845" / "reader_v2.py"

OWNER = "litchi_xlsx::cell_values::source::SourceBackedEditor::publish_multi_commit_to_stream"
MEASURED_PARENT = "litchi_perf_baseline::run_xlsx_cell_values_edit_save"
LIFECYCLE_PARENT = "litchi_perf_baseline::run_xlsx_cell_value_lifecycle_gates"
TOPOLOGY = "litchi_opc::source_backed::SourceBackedPackage::write_topology_to_stream"

ANNOTATION_ENV = {
    "PATH": "/usr/bin:/bin",
    "LC_ALL": "C",
    "LANG": "C",
    "TZ": "UTC",
    "PERL_HASH_SEED": "0",
    "PERL_PERTURB_KEYS": "0",
}

STAR_RE = re.compile(r"^\s*([0-9,]+)\s+\*\s+(.*)$")
EDGE_RE = re.compile(r"^\s*([0-9,]+)\s+>\s+(.*)$")
OPTION_RE = re.compile(r"^(--[^=]+)=(.*)$")


class EvidenceError(ValueError):
    """A missing, malformed, or contradictory retained artifact."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise EvidenceError(message)


def sha256(path: Path) -> str:
    require(path.is_file() and not path.is_symlink(), f"missing regular file: {path}")
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1 << 20), b""):
            digest.update(block)
    return digest.hexdigest()


def read_json(path: Path) -> Any:
    require(path.is_file() and not path.is_symlink(), f"missing JSON file: {path}")
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise EvidenceError(f"cannot read JSON {path}: {error}") from error


def relative(path: Path) -> str:
    try:
        return str(path.relative_to(HERE))
    except ValueError as error:
        raise EvidenceError(f"path escapes packet: {path}") from error


def load_reader() -> Any:
    """Load the retained strict 0845 parser without importing packet code."""

    require(READER_PATH.is_file(), f"missing retained reader: {READER_PATH}")
    spec = importlib.util.spec_from_file_location("retained_callgrind_reader_0845", READER_PATH)
    require(spec is not None and spec.loader is not None,
            f"cannot load retained reader: {READER_PATH}")
    module = importlib.util.module_from_spec(spec)
    # The reader is self-contained, but registering it gives importlib the same
    # normal module semantics used by the retained replay helpers.
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


READER = load_reader()


def validate_plan() -> dict[str, Any]:
    plan = read_json(HERE / "plan.json")
    require(isinstance(plan, dict), "plan is not an object")
    require(plan.get("schema") == "litchi.xlsx-publication-profile-0901.v1",
            "plan schema differs")
    require(plan.get("owner") == OWNER, "plan owner differs")
    require(plan.get("measured_parent") == MEASURED_PARENT,
            "plan measured parent differs")
    require(plan.get("lifecycle_parent") == LIFECYCLE_PARENT,
            "plan lifecycle parent differs")
    require(plan.get("shapes") == ["medium", "dense-sparse"],
            "plan shape order differs")
    require(plan.get("repeats") == 2 and plan.get("samples") == 1 and plan.get("warmup") == 0,
            "plan sample policy differs")
    options = plan.get("options")
    require(isinstance(options, list) and all(isinstance(item, str) for item in options),
            "plan Callgrind options are invalid")
    expected_options = [
        "--tool=callgrind",
        "--collect-atstart=no",
        "--toggle-collect=" + OWNER,
        "--zero-before=" + OWNER,
        "--dump-after=" + OWNER,
    ]
    require(options == expected_options, "plan Callgrind options differ")
    schedule = plan.get("schedule")
    require(isinstance(schedule, list) and schedule, "plan schedule is empty")
    for index, row in enumerate(schedule):
        require(isinstance(row, dict), f"schedule[{index}] is not an object")
        require(row.get("index") == index, f"schedule[{index}] index is not contiguous")
        require(row.get("shape") in plan["shapes"], f"schedule[{index}] shape is unknown")
        require(type(row.get("repeat")) is int and row["repeat"] in range(plan["repeats"]),
                f"schedule[{index}] repeat is invalid")
    require(len(schedule) == plan["repeats"] * len(plan["shapes"]),
            "plan schedule cardinality differs")
    require(type(plan.get("expected_measured_calls")) is int
            and plan["expected_measured_calls"] == 1,
            "plan measured-call expectation differs")
    reader_hash = sha256(READER_PATH)
    plan["_reader_sha256"] = reader_hash
    return plan


def _path_value(command: Sequence[str], option: str) -> str:
    for token in command:
        if token.startswith(option + "="):
            return token.split("=", 1)[1]
    for index, token in enumerate(command[:-1]):
        if token == option:
            return command[index + 1]
    raise EvidenceError(f"command omits {option}")


def _artifact_map(value: Any, folder: Path, label: str) -> dict[str, dict[str, Any]]:
    require(isinstance(value, list), f"{label}: artifacts are not a list")
    result: dict[str, dict[str, Any]] = {}
    for index, item in enumerate(value):
        require(isinstance(item, dict), f"{label}[{index}] is not an object")
        raw_path = item.get("path")
        require(isinstance(raw_path, str), f"{label}[{index}] path is missing")
        path = Path(raw_path)
        require(path.is_file() and not path.is_symlink(),
                f"{label}[{index}] artifact is missing: {path}")
        try:
            path.relative_to(folder)
        except ValueError as error:
            raise EvidenceError(f"{label}[{index}] artifact escapes capture folder: {path}") from error
        name = path.name
        require(name not in result, f"{label}: duplicate artifact {name}")
        digest = item.get("sha256")
        require(isinstance(digest, str) and len(digest) == 64,
                f"{label}[{index}] hash is invalid")
        require(sha256(path) == digest, f"{label}[{index}] hash differs: {path}")
        require(item.get("bytes") == path.stat().st_size,
                f"{label}[{index}] byte count differs: {path}")
        result[name] = {"path": str(path), "bytes": path.stat().st_size, "sha256": digest}
    return result


def validate_capture_metadata(plan: dict[str, Any], row: dict[str, Any], folder: Path) -> dict[str, Any]:
    command_path = folder / "command.json"
    report_path = folder / "report.json"
    command_doc = read_json(command_path)
    report = read_json(report_path)
    require(isinstance(command_doc, dict), f"{relative(command_path)} is not an object")
    argv = command_doc.get("argv")
    require(isinstance(argv, list) and all(isinstance(item, str) for item in argv),
            f"{relative(command_path)} argv is invalid")
    require(command_doc.get("exit_code") == 0, f"{relative(command_path)} workload failed")
    require(command_doc.get("row") == row, f"{relative(command_path)} row differs")
    require(command_doc.get("plan", {}).get("sha256") == sha256(HERE / "plan.json"),
            f"{relative(command_path)} plan binding differs")
    require(argv[:4] == ["/usr/bin/taskset", "-c", str(plan["cpu"]), "/usr/bin/valgrind"],
            f"{relative(command_path)} command prefix differs")
    for token in plan["options"]:
        require(token in argv, f"{relative(command_path)} omits {token}")
    require(_path_value(argv, "--callgrind-out-file") == str(folder / "callgrind.out"),
            f"{relative(command_path)} Callgrind output path differs")
    require(_path_value(argv, "--json") == str(report_path),
            f"{relative(command_path)} report output path differs")
    require(_path_value(argv, "--case") == plan["case"],
            f"{relative(command_path)} case differs")
    require(_path_value(argv, "--xlsx-cell-crud-shape") == row["shape"],
            f"{relative(command_path)} shape differs")
    require(_path_value(argv, "--warmup") == str(plan["warmup"]),
            f"{relative(command_path)} warmup differs")
    require(_path_value(argv, "--samples") == str(plan["samples"]),
            f"{relative(command_path)} samples differs")
    artifacts = _artifact_map(command_doc.get("artifacts"), folder,
                              relative(command_path) + ".artifacts")
    for required in ("callgrind.out", "report.json", "start.json", "stdout.txt", "stderr.txt"):
        require(required in artifacts, f"{relative(command_path)} omits artifact {required}")
    require(isinstance(report, dict), f"{relative(report_path)} is not an object")
    require(report.get("schema_version") == 1, f"{relative(report_path)} schema differs")
    binary = report.get("binary_identity")
    require(isinstance(binary, dict) and binary.get("binary_sha256") == plan["binary"]["sha256"],
            f"{relative(report_path)} binary binding differs")
    configuration = report.get("configuration")
    require(isinstance(configuration, dict), f"{relative(report_path)} configuration is missing")
    require(configuration.get("cases") == [plan["case"]], f"{relative(report_path)} case differs")
    require(configuration.get("xlsx_cell_crud_shapes") == [row["shape"]],
            f"{relative(report_path)} shape differs")
    require(configuration.get("samples_per_case") == plan["samples"]
            and configuration.get("warmup_iterations_per_case") == plan["warmup"],
            f"{relative(report_path)} sample policy differs")
    environment = report.get("environment")
    require(isinstance(environment, dict), f"{relative(report_path)} environment is missing")
    for key in ("git_revision", "git_worktree_dirty", "rustc_version"):
        require(environment.get(key) is None,
                f"{relative(report_path)} expected detached-source {key}=null")
    semantics_path = folder / "semantics.json"
    require(semantics_path.is_file(), f"{relative(semantics_path)} is missing")
    semantics = read_json(semantics_path)
    require(isinstance(semantics, dict) and semantics.get("status") == "pass",
            f"{relative(semantics_path)} semantic gate did not pass")
    return {
        "command": {"file": relative(command_path), "sha256": sha256(command_path)},
        "report": {"file": relative(report_path), "sha256": sha256(report_path)},
        "semantics": {"file": relative(semantics_path), "sha256": sha256(semantics_path)},
        "artifacts": artifacts,
        "binary": binary,
        "environment": environment,
        "configuration": configuration,
    }


def owner_rows(parsed: Mapping[str, Any], owner: str, measured_parent: str,
               lifecycle_parent: str) -> dict[str, Any]:
    names = parsed["names"]
    ids = sorted(function_id for function_id, name in names.items() if name == owner)
    require(ids, f"{parsed['path']}: owner name has no function record")
    incoming: list[dict[str, Any]] = []
    for target_id in ids:
        incoming.extend(READER.incoming_edges(parsed, target_id))
    require(incoming, f"{parsed['path']}: owner has no positive incoming edge")
    measured = [edge for edge in incoming if edge["caller"] == measured_parent]
    lifecycle = [edge for edge in incoming if edge["caller"] == lifecycle_parent]
    known = measured + lifecycle
    unknown = [edge for edge in incoming if edge["caller"] not in {measured_parent, lifecycle_parent}]
    require(not unknown, f"{parsed['path']}: owner has unknown incoming parents: {unknown}")
    require(len(known) == 1, f"{parsed['path']}: owner incoming edge is ambiguous: {known}")
    edge = known[0]
    require(edge["calls"] == 1, f"{parsed['path']}: owner incoming call count is not one")
    role = "measured" if measured else "lifecycle"
    function_id = edge["callee_id"]
    require(function_id in parsed["functions"],
            f"{parsed['path']}: selected owner function record is missing")
    view = READER.function_view(parsed["functions"][function_id], names)
    require(view["self_ir"] + view["direct_children_ir"] == edge["inclusive_ir"],
            f"{parsed['path']}: owner self plus direct Ir does not conserve")
    require(edge["inclusive_ir"] == parsed["header"]["summary_ir"],
            f"{parsed['path']}: owner incoming Ir differs from raw summary")
    return {
        "role": role,
        "owner_ids": ids,
        "incoming_edges": incoming,
        "selected": edge,
        "function": view,
        "validation": {
            "exactly_one_positive_owner_edge": True,
            "known_parent_only": True,
            "one_owner_call": True,
            "owner_ir_matches_summary": True,
            "self_plus_direct_conserves": True,
        },
    }


def numbered_paths(folder: Path) -> list[Path]:
    paths = [path for path in folder.glob("callgrind.out.[0-9]*")
             if path.is_file() and path.suffix[1:].isdigit()]
    paths.sort(key=lambda path: int(path.suffix[1:]))
    require(paths, f"{relative(folder)} has no numbered Callgrind dumps")
    numbers = [int(path.suffix[1:]) for path in paths]
    require(numbers == list(range(1, len(numbers) + 1)),
            f"{relative(folder)} numbered parts are not contiguous: {numbers}")
    return paths


def analyze_terminal(path: Path, last_part: int) -> dict[str, Any]:
    parsed = READER.parse_raw(path, allow_empty=True)
    header = parsed["header"]
    require(header["part"] > last_part, f"{relative(path)} terminal part does not follow numbered dumps")
    require(header["trigger"] == "Program termination", f"{relative(path)} terminal trigger differs")
    require(header["summary_ir"] == 0 and header["totals_ir"] == 0,
            f"{relative(path)} terminal Ir is not zero")
    require(not parsed["functions"], f"{relative(path)} terminal dump has function records")
    return {"file": relative(path), "part": header["part"], "sha256": sha256(path),
            "summary_ir": 0, "totals_ir": 0, "trigger": header["trigger"]}


def display_name(text: str) -> str:
    text = text.split(" [", 1)[0].strip()
    text = re.sub(r"\s+\([0-9,]+x\)$", "", text)
    if text.startswith("???:"):
        return text[4:]
    return text.rsplit(":", 1)[-1] if text.startswith(("./", "/")) else text


def parse_annotation(text: str, selected: str, label: str) -> dict[str, Any]:
    lines = text.splitlines()
    hits: list[tuple[int, int, str]] = []
    for index, line in enumerate(lines):
        match = STAR_RE.match(line)
        if match and display_name(match.group(2)) == selected:
            hits.append((index, int(match.group(1).replace(",", "")), match.group(2)))
    require(len(hits) == 1, f"{label}: expected one annotation row for {selected}, got {hits}")
    row_index, selected_ir, rendered = hits[0]
    direct: list[dict[str, Any]] = []
    for line in lines[row_index + 1:]:
        match = EDGE_RE.match(line)
        if not match:
            break
        rendered_name = match.group(2)
        calls = re.search(r"\(([0-9,]+)x\)$", rendered_name)
        direct.append({
            "name": display_name(rendered_name),
            "rendered": rendered_name,
            "inclusive_ir": int(match.group(1).replace(",", "")),
            "calls": int(calls.group(1).replace(",", "")) if calls else None,
        })
    return {"selected_ir": selected_ir, "rendered": rendered, "direct": direct}


def direct_map(rows: Sequence[Mapping[str, Any]], name_key: str = "name") -> dict[str, int]:
    result: dict[str, int] = {}
    for row in rows:
        name = str(row[name_key])
        result[name] = result.get(name, 0) + int(row["inclusive_ir"])
    return result


def run_annotation(path: Path, inclusive: bool, output: Path, selected: str,
                   expected_raw: int) -> dict[str, Any]:
    command = [
        "/usr/bin/callgrind_annotate", "--auto=no", "--threshold=100",
        "--show-percs=no", "--inclusive=" + ("yes" if inclusive else "no"),
        "--tree=both", str(path),
    ]
    env = dict(ANNOTATION_ENV)
    process = subprocess.run(command, capture_output=True, text=True, env=env)
    require(process.returncode == 0, f"callgrind_annotate failed for {relative(path)}")
    require(process.stderr == "", f"callgrind_annotate emitted stderr for {relative(path)}: {process.stderr}")
    text = process.stdout
    if output.exists():
        require(output.read_text(encoding="utf-8") == text,
                f"annotation replay differs: {relative(output)}")
    else:
        output.write_text(text, encoding="utf-8")
    parsed = parse_annotation(text, selected, relative(output))
    require(parsed["selected_ir"] == expected_raw,
            f"{relative(output)} annotation Ir differs from raw owner")
    direct_ir = sum(item["inclusive_ir"] for item in parsed["direct"])
    return {
        "command": command,
        "environment": dict(ANNOTATION_ENV),
        "file": relative(output),
        "sha256": sha256(output),
        "selected_ir": parsed["selected_ir"],
        "selected_rendered": parsed["rendered"],
        "direct": parsed["direct"],
        "direct_ir": direct_ir,
    }


def annotation_pair(folder: Path, selected_path: Path, owner_function: dict[str, Any]) -> dict[str, Any]:
    inclusive = run_annotation(selected_path, True, folder / "annotation-inclusive.txt",
                                OWNER, owner_function["self_ir"] + owner_function["direct_children_ir"])
    exclusive = run_annotation(selected_path, False, folder / "annotation-self.txt",
                               OWNER, owner_function["self_ir"] + owner_function["direct_children_ir"])
    require(exclusive["direct"] == inclusive["direct"],
            f"{relative(folder)} annotation direct children differ between modes")
    require(exclusive["selected_ir"] + exclusive["direct_ir"] == inclusive["selected_ir"],
            f"{relative(folder)} annotation self plus direct Ir does not conserve")
    raw_map = direct_map(owner_function["direct_children"], "callee")
    annotation_map = direct_map(inclusive["direct"])
    require(raw_map == annotation_map,
            f"{relative(folder)} raw and annotation direct children differ")
    return {
        "inclusive": inclusive,
        "self": exclusive,
        "validation": {
            "inclusive_matches_raw": True,
            "self_direct_matches_inclusive": True,
            "self_plus_direct_conserves": True,
            "raw_direct_children_match_annotation": True,
        },
    }


def named_function_diagnostics(parsed: Mapping[str, Any], names: Sequence[str], parent_id: int | None = None) -> list[dict[str, Any]]:
    """Return exact-name function views, optionally restricted by one caller."""

    result: list[dict[str, Any]] = []
    wanted = set(names)
    for function_id, function_name in sorted(parsed["names"].items()):
        if function_name not in wanted or function_id not in parsed["functions"]:
            continue
        incoming = READER.incoming_edges(parsed, function_id)
        if parent_id is not None:
            incoming = [edge for edge in incoming if edge["caller_id"] == parent_id]
        view = READER.function_view(parsed["functions"][function_id], parsed["names"])
        result.append({
            "function": view,
            "incoming": incoming,
            "incoming_ir": sum(edge["inclusive_ir"] for edge in incoming),
            "incoming_calls": sum(edge["calls"] for edge in incoming),
        })
    return result


def topology_partition(parsed: Mapping[str, Any], owner_id: int) -> dict[str, Any]:
    top_matches = named_function_diagnostics(parsed, [TOPOLOGY], owner_id)
    require(len(top_matches) == 1, f"{parsed['path']}: topology owner edge is ambiguous: {top_matches}")
    top = top_matches[0]
    require(len(top["incoming"]) == 1 and top["incoming"][0]["calls"] == 1,
            f"{parsed['path']}: topology incoming edge is not one call")
    top_view = top["function"]
    require(top["incoming_ir"] == top_view["self_ir"] + top_view["direct_children_ir"],
            f"{parsed['path']}: topology self plus direct Ir does not conserve")
    children = top_view["direct_children"]
    require(children, f"{parsed['path']}: topology has no immediate children")
    direct_total = sum(item["inclusive_ir"] for item in children)
    require(top_view["self_ir"] + direct_total == top["incoming_ir"],
            f"{parsed['path']}: topology immediate-child partition does not conserve")
    # The top-level children are the only disjoint publication partition.  The
    # candidate groups below intentionally classify only those direct rows;
    # nested XML and Deflate rows are separately reported and never added here.
    markers = {
        "xml_audit": ("xml_minifier::audit::", "validate_source_part_replacement_xml"),
        "writer": ("write_changed_overlays", "soapberry_zip::preserve::"),
        "compression": ("zlib_rs::deflate::", "flate2::mem::Compress::compress", "deflate"),
    }
    memberships: dict[str, list[str]] = {key: [] for key in markers}
    for child in children:
        child_name = child["callee"]
        for group, group_markers in markers.items():
            if any(marker in child_name for marker in group_markers):
                memberships[group].append(child_name)
    overlaps = sorted({name for values in memberships.values() for name in values
                       if sum(name in other for other in memberships.values()) > 1})
    groups_present = all(memberships[group] for group in memberships)
    disjoint = not overlaps and groups_present
    return {
        "owner": TOPOLOGY,
        "incoming_from_publication": top["incoming"],
        "function": top_view,
        "immediate_children": children,
        "immediate_child_total_ir": direct_total,
        "candidate_groups": {
            group: {"children": values, "eligible_for_sum": disjoint}
            for group, values in memberships.items()
        },
        "candidate_partition": {
            "disjoint": disjoint,
            "overlapping_children": overlaps,
            "all_groups_present": groups_present,
            "sum_allowed": disjoint,
            "reason": (
                "Only immediate topology children can be summed. XML-audit, writer, "
                "and compression functions below that boundary remain overlapping "
                "diagnostics unless each is an exclusive direct-child group."
            ),
        },
        "validation": {
            "one_topology_incoming_call": True,
            "topology_self_plus_direct_conserves": True,
            "immediate_children_are_disjoint_boundary": True,
        },
    }


def nested_diagnostics(parsed: Mapping[str, Any]) -> dict[str, Any]:
    names = [
        "litchi_opc::source_backed::SourceBackedPackage::prepare_topology_state",
        "litchi_opc::source_backed::SourceBackedPackage::write_changed_overlays_with_appended_inner",
        "litchi_opc::source_backed::validate_source_part_replacement_xml",
        "xml_minifier::audit::verify_source_replacement",
        "soapberry_zip::preserve::PreservationIndex<R>::write_prepared",
        "soapberry_zip::preserve::PreservationIndex<R>::prepare_with",
        "zlib_rs::deflate::deflate",
    ]
    rows = named_function_diagnostics(parsed, names)
    for row in rows:
        view = row["function"]
        require(view["self_ir"] + view["direct_children_ir"]
                == row["incoming_ir"] or row["incoming_ir"] == 0,
                f"{parsed['path']}: nested function conservation failed for {view['name']}")
    return {
        "rows": rows,
        "sum_allowed": False,
        "reason": "Nested owners overlap the immediate topology partition and are diagnostics only.",
    }


def analyze_capture(plan: dict[str, Any], row: dict[str, Any], annotate: bool) -> dict[str, Any]:
    folder = HERE / "capture" / f"{row['index']:02d}-{row['shape']}"
    require(folder.is_dir(), f"missing capture folder: {folder}")
    metadata = validate_capture_metadata(plan, row, folder)
    numbered = numbered_paths(folder)
    parts: list[dict[str, Any]] = []
    measured: list[tuple[Path, dict[str, Any], dict[str, Any]]] = []
    lifecycle_count = 0
    for path in numbered:
        parsed = READER.parse_raw(path)
        header = parsed["header"]
        part = int(path.suffix[1:])
        require(header["part"] == part, f"{relative(path)} part differs from suffix")
        require(header["trigger"] == "--dump-after=" + OWNER,
                f"{relative(path)} trigger differs")
        owner = owner_rows(parsed, OWNER, MEASURED_PARENT, LIFECYCLE_PARENT)
        owner["path"] = relative(path)
        owner["sha256"] = sha256(path)
        owner["part"] = part
        owner["summary_ir"] = header["summary_ir"]
        if owner["role"] == "measured":
            measured.append((path, parsed, owner))
        else:
            lifecycle_count += 1
        parts.append({
            "file": relative(path),
            "part": part,
            "sha256": owner["sha256"],
            "header": header,
            "owner": owner,
        })
    require(len(measured) == plan["expected_measured_calls"],
            f"{relative(folder)} measured owner count differs: {len(measured)}")
    terminal = analyze_terminal(folder / "callgrind.out", int(numbered[-1].suffix[1:]))
    selected_path, selected_parsed, selected_owner = measured[0]
    topology = topology_partition(selected_parsed, selected_owner["selected"]["callee_id"])
    nested = nested_diagnostics(selected_parsed)
    annotations = annotation_pair(folder, selected_path, selected_owner["function"]) if annotate else None
    return {
        "index": row["index"],
        "repeat": row["repeat"],
        "shape": row["shape"],
        "metadata": metadata,
        "parts": parts,
        "terminal": terminal,
        "classification": {
            "measured_count": len(measured),
            "lifecycle_count": lifecycle_count,
            "measured_part": selected_owner["part"],
            "measured_file": selected_owner["path"],
            "lifecycle_prediction": plan.get("historical_expected_lifecycle_calls"),
        },
        "measured": selected_owner,
        "annotations": annotations,
        "topology": topology,
        "nested_diagnostics": nested,
        "validation": {
            "all_numbered_dumps_retained": True,
            "all_numbered_dumps_have_known_owner_parent": True,
            "one_measured_owner_call": True,
            "lifecycle_calls_retained_separately": True,
        },
    }


def analyze(output: Path, annotate: bool, first_only: bool) -> dict[str, Any]:
    plan = validate_plan()
    rows = plan["schedule"][:1] if first_only else plan["schedule"]
    captures = [analyze_capture(plan, row, annotate) for row in rows]
    measured = [item for item in captures if item["classification"]["measured_count"] == 1]
    require(len(measured) == len(captures), "capture matrix has an invalid measured count")
    result = {
        "schema": "litchi.xlsx-publication-profile-analysis-0901.v1",
        "status": "pass",
        "mode": "first-child-admission" if first_only else "full",
        "plan": {"file": relative(HERE / "plan.json"), "sha256": sha256(HERE / "plan.json")},
        "retained_reader": {
            "file": str(READER_PATH.relative_to(HERE.parents[3])),
            "sha256": plan["_reader_sha256"],
        },
        "owner": OWNER,
        "measured_parent": MEASURED_PARENT,
        "lifecycle_parent": LIFECYCLE_PARENT,
        "topology_owner": TOPOLOGY,
        "captures": captures,
        "scope": plan.get("scope"),
        "limitations": [
            "Callgrind Ir is a guest-instruction attribution mechanism, not native latency, hardware cycles, allocations, or RSS.",
            "Lifecycle owner calls are classified and retained separately from the one measured runner call.",
            "Only immediate children of the selected owner are a disjoint partition; nested XML, writer, and compression rows are not added.",
        ],
    }
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(json.dumps(result, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    return result


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", type=Path, default=HERE / "analysis.json")
    parser.add_argument("--first-child", action="store_true",
                        help="analyze only schedule index 0 for admission")
    parser.add_argument("--no-annotate", action="store_true",
                        help="skip callgrind_annotate; retain raw-only analysis")
    args = parser.parse_args()
    try:
        report = analyze(args.output, annotate=not args.no_annotate, first_only=args.first_child)
    except (EvidenceError, OSError, subprocess.SubprocessError) as error:
        print(f"analyze.py: error: {error}", file=sys.stderr)
        return 2
    print(f"0901 Callgrind analysis passed: {args.output} ({len(report['captures'])} capture(s))")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
