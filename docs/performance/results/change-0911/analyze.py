#!/usr/bin/env python3
"""Offline, fail-closed reducer for the 0911 v3 snapshot profiles.

The retained 0845 reader owns the Callgrind grammar.  This adapter binds the
v3 plan, child receipts, workload reports, and the two expected dynamic owner
contexts.  It never builds, runs a workload, invokes Valgrind, or calls an
external command.
"""

from __future__ import annotations

import argparse
import hashlib
import importlib.util
import json
from pathlib import Path
import sys
from typing import Any, Mapping


HERE = Path(__file__).resolve().parent
REPO = HERE.parents[3]
PLAN_PATH = HERE / "plan-v3.json"
READER_PATH = HERE.parent / "change-0845" / "reader_v2.py"
CAPTURE_PATH = HERE / "capture_v3.py"
OWNER = "litchi_pptx::package::model::Package::opened_presentation_transaction"
PARENT = "litchi_perf_baseline::pptx_owned_benchmark::timed_iteration"
SETUP_PARENT = "litchi_perf_baseline::build_pptx_source_edit_corpus"
LABELS = ("materialized-a", "owned-a", "owned-b", "materialized-b")
KINDS = {
    "materialized-a": "materialized",
    "owned-a": "owned",
    "owned-b": "owned",
    "materialized-b": "materialized",
}
EXPECTED_OPTIONS = (
    "--tool=callgrind",
    "--collect-atstart=no",
    "--toggle-collect=" + OWNER,
    "--zero-before=" + OWNER,
    "--dump-after=" + OWNER,
)
EXPECTED_CPU = 12
EXPECTED_ENVIRONMENT = {"PATH": "/usr/bin:/bin", "LC_ALL": "C", "LANG": "C", "TZ": "UTC"}
EXPECTED_CWD = "/tmp/litchi-0911-run-v3"
EXPECTED_SCHEMA = "litchi.pptx-snapshot-instruction-profile-0911.v3"
EXPECTED_BINARY_SHA256 = "f4ad7ee158061cc80354c3c8674c9fe6a8100ed76a0c8f743653efcb06bb2e8d"
EXPECTED_SOURCE_SHA256 = "61b2b99083ca27ebd37955db600955e3f41289b93dba71951983164239eff757"
EXPECTED_OUTPUT_SHA256 = "23d6d7b8dd433ff453307ee56485efadc34d6b18b455a6858d9f565c4b1b6cd9"
EXPECTED_SOURCE_BYTES = 17_017_139
EXPECTED_OUTPUT_BYTES = 17_017_145
EXPECTED_MEMBER_COUNT = 445
EXPECTED_ENTRY_COUNT = 229
TARGET_NAMES = (
    "litchi_pptx::package::model::Package::opened_presentation_with_limits",
    "litchi_pptx::opened::model::capture_internal",
    "litchi_pptx::opened::model::package_fingerprint_with_memo",
    "litchi_pptx::opened::transaction::Transaction::new",
)


class EvidenceError(ValueError):
    """A retained artifact is missing, malformed, or contradictory."""


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


def descriptor(path: Path) -> dict[str, Any]:
    path = path.resolve()
    return {"path": str(path), "bytes": path.stat().st_size, "sha256": sha256(path)}


def read_json(path: Path) -> Any:
    require(path.is_file() and not path.is_symlink(), f"missing JSON: {path}")
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise EvidenceError(f"cannot read JSON {path}: {error}") from error


def checked_descriptor(value: Any, label: str) -> dict[str, Any]:
    require(isinstance(value, Mapping), f"{label} is not a descriptor")
    raw = value.get("path")
    require(isinstance(raw, str), f"{label}.path is missing")
    actual = descriptor(Path(raw))
    require(value.get("bytes") == actual["bytes"], f"{label}.bytes differs")
    require(value.get("sha256") == actual["sha256"], f"{label}.sha256 differs")
    return actual


def packet_relative(path: Path) -> str:
    try:
        return str(path.resolve().relative_to(HERE.resolve()))
    except ValueError as error:
        raise EvidenceError(f"path escapes 0911 packet: {path}") from error


def load_reader() -> Any:
    require(READER_PATH.is_file(), f"missing retained reader: {READER_PATH}")
    spec = importlib.util.spec_from_file_location("retained_callgrind_reader_0845_0911", READER_PATH)
    require(spec is not None and spec.loader is not None, "cannot load retained reader")
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


READER = load_reader()


def load_capture_helper() -> Any:
    require(CAPTURE_PATH.is_file(), f"missing capture helper: {CAPTURE_PATH}")
    spec = importlib.util.spec_from_file_location("retained_capture_helper_0911_v3", CAPTURE_PATH)
    require(spec is not None and spec.loader is not None, "cannot load capture helper")
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


CAPTURE = load_capture_helper()


def normalize_command(value: str) -> str:
    return " ".join(value.split())


def validate_plan(plan: Mapping[str, Any]) -> dict[str, Any]:
    require(plan.get("schema") == EXPECTED_SCHEMA, "plan schema differs")
    require(plan.get("owner") == OWNER and plan.get("parent") == PARENT,
            "plan owner or parent differs")
    require(plan.get("samples") == 1 and plan.get("warmup") == 0,
            "plan sample or warmup differs")
    require(plan.get("cpu") == EXPECTED_CPU, "plan CPU differs")
    require(plan.get("cwd") == EXPECTED_CWD, "plan cwd differs")
    require(plan.get("environment") == EXPECTED_ENVIRONMENT,
            "plan environment differs")
    require(plan.get("options") == list(EXPECTED_OPTIONS), "plan Callgrind options differ")
    require(plan.get("schedule") == [
        {"index": 0, "label": "materialized-a", "kind": "materialized"},
        {"index": 1, "label": "owned-a", "kind": "owned"},
        {"index": 2, "label": "owned-b", "kind": "owned"},
        {"index": 3, "label": "materialized-b", "kind": "materialized"},
    ], "plan schedule differs")
    binary = checked_descriptor(plan.get("binary"), "plan.binary")
    require(binary["sha256"] == EXPECTED_BINARY_SHA256, "plan binary hash differs")
    helper = checked_descriptor(plan.get("helper"), "plan.helper")
    require(helper == descriptor(CAPTURE_PATH), "plan helper differs from retained capture helper")
    for key in ("inputs", "tools"):
        rows = plan.get(key)
        require(isinstance(rows, list) and rows, f"plan.{key} is missing")
        for index, row in enumerate(rows):
            checked_descriptor(row, f"plan.{key}[{index}]")
    expected = plan.get("expected_positive_callers")
    require(expected == {"setup": SETUP_PARENT, "measured": PARENT},
            "plan expected owner callers differ")
    return {
        "schema": plan["schema"],
        "binary": binary,
        "owner": OWNER,
        "parent": PARENT,
        "expected_positive_callers": dict(expected),
        "options": list(plan["options"]),
        "samples": plan["samples"],
        "warmup": plan["warmup"],
        "cpu": EXPECTED_CPU,
        "cwd": EXPECTED_CWD,
        "environment": dict(EXPECTED_ENVIRONMENT),
        "binary_record": dict(plan["binary"]),
        "helper_record": dict(plan["helper"]),
        "schedule": list(plan["schedule"]),
        "plan": descriptor(PLAN_PATH),
    }


def command_binary_token(argv: list[str]) -> str | None:
    for value in argv:
        if value.endswith("/binaries/native"):
            return value
    return None


def validate_command(folder: Path, plan: Mapping[str, Any], row: Mapping[str, Any]) -> dict[str, Any]:
    command_path = folder / "command.json"
    start_path = folder / "start.json"
    command = read_json(command_path)
    start = read_json(start_path)
    require(isinstance(command, Mapping) and isinstance(start, Mapping),
            f"{packet_relative(folder)} receipts are malformed")
    require(command.get("exit_code") == 0, f"{packet_relative(folder)} command failed")
    argv = command.get("argv")
    require(isinstance(argv, list) and all(isinstance(item, str) for item in argv),
            f"{packet_relative(folder)} argv is malformed")
    expected_binary = str(Path(plan["binary_record"]["path"]).resolve())
    expected_argv = [
        "/usr/bin/taskset", "-c", str(plan["cpu"]), "/usr/bin/valgrind",
        *plan["options"], "--callgrind-out-file=" + str((folder / "callgrind.out").resolve()),
        expected_binary, str(row["kind"]), "0", "1", str((folder / "artifacts").resolve()),
    ]
    require(argv == expected_argv, f"{packet_relative(folder)} argv differs from frozen plan")
    require(command.get("binary") == plan["binary_record"],
            f"{packet_relative(folder)} binary receipt differs")
    require(command.get("helper") == plan["helper_record"],
            f"{packet_relative(folder)} helper receipt differs")
    require(command.get("plan") == plan["plan"],
            f"{packet_relative(folder)} plan receipt differs")
    require(command.get("environment") == plan["environment"],
            f"{packet_relative(folder)} environment differs")
    require(command.get("cwd") == plan["cwd"],
            f"{packet_relative(folder)} cwd differs")
    require(command.get("row") == dict(row), f"{packet_relative(folder)} schedule row differs")
    require(start.get("argv") == command.get("argv"),
            f"{packet_relative(folder)} start/command argv differs")
    require(start.get("binary") == command.get("binary")
            and start.get("helper") == command.get("helper")
            and start.get("plan") == command.get("plan")
            and start.get("cwd") == command.get("cwd")
            and start.get("environment") == command.get("environment")
            and start.get("row") == command.get("row"),
            f"{packet_relative(folder)} start/command custody differs")
    start_argv = start.get("argv")
    require(start_argv == expected_argv, f"{packet_relative(folder)} start argv differs")
    require(start.get("plan") == plan["plan"] and start.get("binary") == plan["binary_record"],
            f"{packet_relative(folder)} start custody differs")
    require(start.get("helper") == plan["helper_record"]
            and start.get("cwd") == plan["cwd"]
            and start.get("environment") == plan["environment"]
            and start.get("row") == dict(row),
            f"{packet_relative(folder)} start receipt differs")
    artifacts = command.get("artifacts")
    require(isinstance(artifacts, list) and artifacts, f"{packet_relative(folder)} artifacts missing")
    actual_artifacts = []
    for index, item in enumerate(artifacts):
        actual_artifacts.append(checked_descriptor(item, f"{packet_relative(folder)}.artifacts[{index}]"))
    return {
        "command": descriptor(command_path),
        "start": descriptor(start_path),
        "exit_code": command["exit_code"],
        "argv": list(argv),
        "expected_target_command": normalize_command(
            " ".join(expected_argv[expected_argv.index(expected_binary):])),
        "artifacts": actual_artifacts,
        "validation": {
            "exact_argv_matches": True,
            "binary_matches": True,
            "options_match": True,
            "environment_matches": True,
            "cwd_matches": True,
            "artifact_descriptors_match": True,
        },
    }


def validate_workload(folder: Path, plan: Mapping[str, Any], row: Mapping[str, Any]) -> dict[str, Any]:
    report_path = folder / "artifacts" / "report.json"
    semantics_path = folder / "semantics.json"
    report = read_json(report_path)
    semantics = read_json(semantics_path)
    require(isinstance(report, Mapping) and isinstance(semantics, Mapping),
            f"{packet_relative(folder)} workload receipts are malformed")
    expected_semantics = CAPTURE.validate_semantics(read_json(PLAN_PATH), dict(row), folder)
    require(semantics == expected_semantics, f"{packet_relative(folder)} semantics receipt differs from retained verifier")
    require(semantics.get("status") == "pass", f"{packet_relative(folder)} semantic gate failed")
    require(report.get("schema") == "pptx-owned-input-benchmark-v1",
            f"{packet_relative(folder)} report schema differs")
    require(report.get("status") == "completed" and report.get("route") == row["kind"],
            f"{packet_relative(folder)} report route/status differs")
    require(report.get("warmup_iterations") == 0 and report.get("measured_samples") == 1,
            f"{packet_relative(folder)} sample envelope differs")
    corpus = report.get("corpus")
    require(isinstance(corpus, Mapping), f"{packet_relative(folder)} corpus missing")
    require(corpus.get("archive_bytes") == EXPECTED_SOURCE_BYTES
            and corpus.get("archive_sha256") == EXPECTED_SOURCE_SHA256
            and corpus.get("archive_member_count") == EXPECTED_MEMBER_COUNT
            and corpus.get("entry_count") == EXPECTED_ENTRY_COUNT,
            f"{packet_relative(folder)} corpus identity differs")
    source = report.get("source")
    output = report.get("output")
    require(isinstance(source, Mapping) and isinstance(output, Mapping),
            f"{packet_relative(folder)} source/output missing")
    require(source.get("bytes") == EXPECTED_SOURCE_BYTES
            and source.get("sha256") == EXPECTED_SOURCE_SHA256,
            f"{packet_relative(folder)} source identity differs")
    require(output.get("bytes") == EXPECTED_OUTPUT_BYTES
            and output.get("sha256") == EXPECTED_OUTPUT_SHA256,
            f"{packet_relative(folder)} output identity differs")
    sample_rows = report.get("samples")
    require(isinstance(sample_rows, list) and len(sample_rows) == 1,
            f"{packet_relative(folder)} sample list differs")
    sample = sample_rows[0]
    require(isinstance(sample, Mapping) and sample.get("index") == 0,
            f"{packet_relative(folder)} sample row differs")
    gates = report.get("gates")
    require(isinstance(gates, Mapping) and gates and all(value is True for value in gates.values()),
            f"{packet_relative(folder)} workload gates are not all true")
    for name in ("source", "output"):
        value = report[name]
        artifact_path = folder / "artifacts" / str(value.get("file"))
        actual = descriptor(artifact_path)
        require(actual["bytes"] == value.get("bytes") and actual["sha256"] == value.get("sha256"),
                f"{packet_relative(folder)} {name} artifact differs")
    return {
        "report": descriptor(report_path),
        "semantics": descriptor(semantics_path),
        "route": row["kind"],
        "corpus": dict(corpus),
        "source": dict(source),
        "output": dict(output),
        "sample": dict(sample),
        "gates": dict(gates),
        "validation": {
            "semantic_status_pass": True,
            "corpus_identity": True,
            "source_identity": True,
            "output_identity": True,
            "source_output_artifacts_match": True,
            "workload_gates_all_true": True,
        },
    }


def direct_rows(parsed: Mapping[str, Any], owner_id: int) -> dict[str, Any]:
    view = READER.function_view(parsed["functions"][owner_id], parsed["names"])
    rows = view["direct_children"]
    return {
        "rows": rows,
        "direct_children_ir": view["direct_children_ir"],
        "row_count": len(rows),
        "inclusive_nonadditive": False,
        "interpretation": "Immediate direct-child inclusive rows form a disjoint partition of owner direct cost; nested descendants are not added.",
    }


def target_rows(parsed: Mapping[str, Any]) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for target in TARGET_NAMES:
        ids = sorted(fid for fid, function in parsed["functions"].items()
                     if function.get("name") == target)
        for function_id in ids:
            view = READER.function_view(parsed["functions"][function_id], parsed["names"])
            rows.append({
                "name": target,
                "id": function_id,
                "file": view["file"],
                "self_ir": view["self_ir"],
                "direct_children_ir": view["direct_children_ir"],
                "direct_children": view["direct_children"],
                "incoming": READER.incoming_edges(parsed, function_id),
                "inclusive_nonadditive": True,
                "interpretation": "Nested function and incoming-edge values are inclusive diagnostics; do not add them to the owner partition.",
            })
    return rows


def parse_positive(path: Path, expected_target_command: str, label: str) -> dict[str, Any]:
    parsed = READER.parse_raw(path)
    header = parsed["header"]
    require(isinstance(header.get("command"), str), f"{label}: raw command is missing")
    require(normalize_command(header["command"]) == expected_target_command,
            f"{label}: raw command differs from command receipt")
    require(header["events"] == ["Ir"], f"{label}: event set differs")
    require(header["summary_ir"] > 0 and header["totals_ir"] == header["summary_ir"],
            f"{label}: positive summary/totals mismatch")
    require(header["trigger"] == "--dump-after=" + OWNER,
            f"{label}: positive trigger differs")
    owner_ids = sorted(fid for fid, function in parsed["functions"].items()
                       if function.get("name") == OWNER)
    require(len(owner_ids) == 1, f"{label}: owner resolves to {owner_ids}")
    owner_id = owner_ids[0]
    incoming = READER.incoming_edges(parsed, owner_id)
    require(len(incoming) == 1, f"{label}: expected one positive owner edge, got {incoming}")
    edge = incoming[0]
    require(edge["calls"] == 1 and edge["inclusive_ir"] == header["summary_ir"],
            f"{label}: owner edge does not match summary")
    require(edge["caller"] in (SETUP_PARENT, PARENT),
            f"{label}: unexpected owner caller {edge['caller']!r}")
    view = READER.function_view(parsed["functions"][owner_id], parsed["names"])
    all_self = sum(function["self_ir"] for function in parsed["functions"].values())
    require(view["self_ir"] + view["direct_children_ir"] == header["summary_ir"],
            f"{label}: owner self/direct partition does not conserve")
    require(all_self == header["summary_ir"], f"{label}: all function self costs do not conserve")
    return {
        "file": packet_relative(path),
        "sha256": sha256(path),
        "bytes": path.stat().st_size,
        "part": header["part"],
        "header": header,
        "owner": {
            "id": owner_id,
            "name": OWNER,
            "caller": edge["caller"],
            "incoming": incoming,
            "self_ir": view["self_ir"],
            "direct_children_ir": view["direct_children_ir"],
            "summary_ir": header["summary_ir"],
            "all_function_self_ir": all_self,
            "direct_children": direct_rows(parsed, owner_id),
            "inclusive_nonadditive": False,
        },
        "nested_targets": target_rows(parsed),
        "exclusive_self_rows": sorted(({
            "id": function["id"],
            "name": function.get("name") or "<unnamed>",
            "file": function.get("file", ""),
            "self_ir": function["self_ir"],
            "records": function["records"],
        } for function in parsed["functions"].values() if function["self_ir"] > 0),
        key=lambda row: (-row["self_ir"], row["name"], row["id"])),
        "exclusive_focus_rows": [row for row in sorted(({
            "id": function["id"],
            "name": function.get("name") or "<unnamed>",
            "file": function.get("file", ""),
            "self_ir": function["self_ir"],
            "records": function["records"],
        } for function in parsed["functions"].values()
            if function["self_ir"] > 0
            and any(token in (function.get("name") or "").casefold()
                    for token in ("sha", "inflate", "deflate"))),
            key=lambda row: (-row["self_ir"], row["name"], row["id"]))],
        "parser": parsed["statistics"],
        "validation": {
            "exact_ir_event": True,
            "positive_summary_equals_totals": True,
            "exact_owner_trigger": True,
            "one_owner_function": True,
            "one_positive_owner_edge": True,
            "owner_edge_matches_summary": True,
            "expected_caller": True,
            "owner_partition_conserved": True,
            "all_function_self_conserved": True,
        },
    }


def parse_terminal(path: Path, expected_part: int, expected_target_command: str,
                   label: str) -> dict[str, Any]:
    parsed = READER.parse_raw(path, allow_empty=True)
    header = parsed["header"]
    require(isinstance(header.get("command"), str), f"{label}: terminal command is missing")
    require(normalize_command(header["command"]) == expected_target_command,
            f"{label}: terminal command differs from command receipt")
    require(header["part"] == expected_part, f"{label}: terminal part differs")
    require(header["events"] == ["Ir"] and header["trigger"] == "Program termination",
            f"{label}: terminal header differs")
    require(header["summary_ir"] == 0 and header["totals_ir"] == 0 and not parsed["functions"],
            f"{label}: terminal dump is not empty zero Ir")
    return {
        "file": packet_relative(path),
        "sha256": sha256(path),
        "bytes": path.stat().st_size,
        "header": header,
        "validation": {"zero_summary": True, "zero_totals": True, "empty_function_graph": True},
    }


def parse_leg(plan: Mapping[str, Any], row: Mapping[str, Any]) -> dict[str, Any]:
    label = str(row["label"])
    folder = HERE / "capture-v3" / label
    require(folder.is_dir() and not folder.is_symlink(), f"missing capture folder: {folder}")
    command = validate_command(folder, plan, row)
    workload = validate_workload(folder, plan, row)
    all_callgrind = sorted(folder.glob("callgrind.out*"))
    require(all(path.name == "callgrind.out"
                or (path.name.startswith("callgrind.out.")
                    and path.name.rsplit(".", 1)[-1].isdigit())
                for path in all_callgrind),
            f"{label}: unknown Callgrind output file present")
    require(folder / "callgrind.out" in all_callgrind,
            f"{label}: terminal Callgrind output is missing")
    numbered = sorted(path for path in all_callgrind if path.name != "callgrind.out")
    require(len(numbered) == 2, f"{label}: expected exactly two positive Callgrind parts")
    require(len(all_callgrind) == 3, f"{label}: unexpected number of Callgrind outputs")
    parts = [parse_positive(path, command["expected_target_command"],
                            f"{label}:{path.name}") for path in numbered]
    callers = [part["owner"]["caller"] for part in parts]
    require(sorted(callers) == sorted((SETUP_PARENT, PARENT)),
            f"{label}: expected setup and measured callers exactly once, got {callers}")
    by_context = {part["owner"]["caller"]: part for part in parts}
    terminal = parse_terminal(folder / "callgrind.out", 3,
                              command["expected_target_command"], label)
    return {
        "label": label,
        "kind": row["kind"],
        "command": command,
        "workload": workload,
        "setup": by_context[SETUP_PARENT],
        "measured": by_context[PARENT],
        "terminal": terminal,
        "validation": {
            "exactly_two_positive_parts": True,
            "setup_context_once": True,
            "measured_context_once": True,
            "all_positive_owner_edges_accounted": True,
            "terminal_zero": True,
            "nested_rows_nonadditive": True,
        },
    }


def reduce_packet(first_child: bool) -> dict[str, Any]:
    raw_plan = read_json(PLAN_PATH)
    require(isinstance(raw_plan, Mapping), "plan is not an object")
    plan = validate_plan(raw_plan)
    rows = [dict(item) for item in raw_plan["schedule"]]
    if first_child:
        rows = rows[:1]
    legs = [parse_leg(plan, row) for row in rows]
    measured = [{
        "label": leg["label"],
        "kind": leg["kind"],
        "summary_ir": leg["measured"]["owner"]["summary_ir"],
        "self_ir": leg["measured"]["owner"]["self_ir"],
        "direct_children_ir": leg["measured"]["owner"]["direct_children_ir"],
        "all_function_self_ir": leg["measured"]["owner"]["all_function_self_ir"],
        "raw_file": leg["measured"]["file"],
        "raw_sha256": leg["measured"]["sha256"],
    } for leg in legs]
    setup = [{
        "label": leg["label"],
        "summary_ir": leg["setup"]["owner"]["summary_ir"],
        "self_ir": leg["setup"]["owner"]["self_ir"],
        "direct_children_ir": leg["setup"]["owner"]["direct_children_ir"],
        "all_function_self_ir": leg["setup"]["owner"]["all_function_self_ir"],
        "raw_file": leg["setup"]["file"],
        "raw_sha256": leg["setup"]["sha256"],
    } for leg in legs]
    require(all(row["all_function_self_ir"] == row["summary_ir"] for row in measured),
            "measured function conservation changed")
    require(all(row["all_function_self_ir"] == row["summary_ir"] for row in setup),
            "setup function conservation changed")
    result = {
        "schema": "litchi.pptx-snapshot-callgrind-analysis-0911.v3",
        "status": "pass",
        "mode": "first-child" if first_child else "full-reduction",
        "plan": plan,
        "owner": OWNER,
        "measured_parent": PARENT,
        "setup_parent": SETUP_PARENT,
        "legs": legs,
        "measured_rows": measured,
        "setup_rows": setup,
        "interpretation": {
            "measured_rows": "Only the exact timed_iteration owner context enters route comparison.",
            "setup_rows": "Corpus-builder owner context is retained as custody/scope diagnostic and is excluded from route comparison.",
            "exclusive_partition": "For each raw part, owner self Ir plus immediate direct-child inclusive Ir equals that part summary; all function self Ir also equals the summary.",
            "nested_rows": "Named nested function rows and incoming edges are inclusive diagnostics and explicitly non-additive.",
            "ir_scope": "Guest Ir only; no elapsed-time, allocation, RSS, physical-I/O, or speedup inference.",
        },
        "checks": {
            "plan_and_custody": True,
            "all_workload_gates": True,
            "exact_two_contexts_per_leg": True,
            "all_positive_owner_edges_accounted": True,
            "owner_partition_conserved": True,
            "all_function_self_conserved": True,
            "terminal_dumps_zero": True,
            "setup_excluded_from_measured_comparison": True,
        },
    }
    if not first_child:
        require(len(legs) == 4, "full reduction requires four legs")
        result["checks"]["four_legs"] = True
        result["route_rows"] = {
            "materialized": [row for row in measured if row["kind"] == "materialized"],
            "owned": [row for row in measured if row["kind"] == "owned"],
        }
    return result


def write_create_only(path: Path, value: Mapping[str, Any]) -> None:
    require(not path.exists(), f"refusing to overwrite existing output: {path}")
    with path.open("x", encoding="utf-8") as stream:
        json.dump(value, stream, indent=2, sort_keys=True)
        stream.write("\n")


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--first-child", action="store_true",
                        help="reduce only materialized-a for v3 admission")
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    try:
        result = reduce_packet(args.first_child)
        if args.first_child:
            result["schema"] = "litchi.pptx-snapshot-first-child-admission-0911.v3"
            result["admission"] = {
                "status": "pass",
                "first_child": "materialized-a",
                "required_before_remaining_legs": True,
            }
        write_create_only(args.output, result)
    except EvidenceError as error:
        print(f"analysis rejected: {error}", file=sys.stderr)
        return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
