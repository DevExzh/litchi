#!/usr/bin/env python3
"""Validate and attribute the matched 0554 N name-handoff profiles.

This is an evidence-only Callgrind analyzer.  It reuses the immutable raw
Callgrind parser from change-0536, but owns the 0554 plan, stage paths, job
matrix, and report schema.  The selected owner edge and positive benchmark
ancestry classify a dump; dump numbering is checked only for continuity and
the historical CFB setup position.  Missing or inlined helpers are reported
explicitly and are never converted into eliminated work.

Callgrind Ir and call metadata are mechanism evidence.  Native latency,
allocation, RSS, and admission decisions belong to independent analyzers.
"""

from __future__ import annotations

import argparse
import importlib.util
import json
from pathlib import Path
import re
import sys
from typing import Any

sys.dont_write_bytecode = True


HERE = Path(__file__).resolve().parent
PLAN_PATH = HERE / "plan.json"
RUN_PATH = HERE / "run.py"
STAGES = ("baseline", "candidate")

RAW_HELPER_PATH = HERE.parent / "change-0536" / "analyze_profiles.py"

XLS_OWNER = (
    "litchi_xls::workbook::source::SourceBackedWorkbook::from_read_at_with_limits"
)
CFB_OWNER = "litchi_cfb::file::OleFile<R>::open"
XLS_RUNNER = "litchi_perf_baseline::run_xls_owned_source_case"
CFB_RUNNER = "litchi_perf_baseline::run_cfb_open"
XLS_OWNER_WRAPPER = "litchi_xls::workbook::source::SourceBackedWorkbook::from_read_at"
CFB_SETUP_CALLERS = (
    "litchi_perf_baseline::build_cfb_corpus",
    "litchi_perf_baseline::run::{{closure}}",
)
MAX_ANCESTRY_DEPTH = 8

PROFILE_FLAGS = (
    "--dump-instr=yes",
    "--dump-line=no",
    "--compress-pos=no",
    "--collect-jumps=yes",
)
CFB_SHAPES = ("tiny", "many-small", "few-large")
XLS_CASE = "xls_owned_source_open_one_cell"
CFB_CASE = "cfb_open"

# The synthetic CFB rows contain root plus the listed stream count.  The XLS
# number is the retained fixed-corpus owner observation.  These values are
# predictions for direct parse_directory_entry -> decode_utf16le calls, not
# counts inferred from the decoder's outgoing allocator edges.
BASELINE_DECODER_CALLS = {
    "cfb-tiny": 4,
    "cfb-many-small": 257,
    "cfb-few-large": 5,
    "xls-owned": 12,
}

TARGET_MARKERS: dict[str, tuple[str, ...]] = {
    "validated_directory_entries": ("::validated_directory_entries",),
    "parse_validated_directory_entry": ("::parse_validated_directory_entry",),
    "directory_name_data": ("directory_name::directory_name_data",),
    "parse_directory_entry": ("::parse_directory_entry",),
    "decode_utf16le": ("::decode_utf16le",),
    "format_clsid": ("::format_clsid",),
    "build_storage_tree_iterative": ("::build_storage_tree_iterative",),
    "name_data_drop": (
        "drop_in_place",
        "DirectoryNameData",
    ),
}

PART_RE = re.compile(r"^part:\s*(\d+)\s*$")
TRIGGER_RE = re.compile(r"^desc:\s+Trigger:\s+(.*)$")
SUMMARY_RE = re.compile(r"^summary:\s*(.*)$")


class EvidenceError(ValueError):
    """A missing, malformed, or contradictory evidence artifact."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise EvidenceError(message)


def sha256(path: Path) -> str:
    try:
        import hashlib

        digest = hashlib.sha256()
        with path.open("rb") as stream:
            for block in iter(lambda: stream.read(1024 * 1024), b""):
                digest.update(block)
        return digest.hexdigest()
    except OSError as error:
        raise EvidenceError(f"cannot hash {path}: {error}") from error


def read_text(path: Path) -> str:
    try:
        return path.read_text(encoding="utf-8", errors="replace")
    except OSError as error:
        raise EvidenceError(f"cannot read {path}: {error}") from error


def read_json(path: Path) -> Any:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise EvidenceError(f"cannot read JSON {path}: {error}") from error


def relative(path: Path) -> str:
    try:
        return path.relative_to(HERE).as_posix()
    except ValueError as error:
        raise EvidenceError(f"path is outside evidence directory: {path}") from error


def load_raw_helper() -> Any:
    require(RAW_HELPER_PATH.is_file(), f"missing immutable parser {RAW_HELPER_PATH}")
    spec = importlib.util.spec_from_file_location(
        "litchi_0554_immutable_raw_profile_parser", RAW_HELPER_PATH
    )
    require(spec is not None and spec.loader is not None,
            f"cannot load immutable parser {RAW_HELPER_PATH}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    # The parser uses HERE only to render errors and relative artifact names.
    # Rebinding this imported module variable does not modify the retained
    # source and avoids importing any old campaign plan or verifier constants.
    module.HERE = HERE
    return module


RAW = load_raw_helper()


def plan_data() -> dict[str, Any]:
    plan = read_json(PLAN_PATH)
    require(isinstance(plan, dict), "plan is not an object")
    require(plan.get("schema") == "ole2_name_handoff_0554_plan_v1",
            "plan schema differs")
    require(plan.get("scope") ==
            "Matched private CFB validated-name ownership handoff N; NF field seed excluded",
            "plan scope differs")
    require(plan.get("cpu") == 2, "plan CPU differs")
    require(plan.get("candidate_files") == ["crates/litchi-cfb/src/file.rs"],
            "candidate file scope differs")
    require(plan.get("groups", {}).get("cfb", {}).get("cases") == [CFB_CASE],
            "CFB case matrix differs")
    require(tuple(plan.get("groups", {}).get("cfb", {}).get("shapes", ())) == CFB_SHAPES,
            "CFB shape matrix differs")
    require(plan.get("groups", {}).get("cfb", {}).get("payload") == "incompressible",
            "CFB payload differs")
    require(XLS_CASE in plan.get("groups", {}).get("xls", {}).get("cases", []),
            "XLS owner case is absent")

    profile = plan.get("profile")
    require(isinstance(profile, dict), "profile plan is missing")
    require(profile.get("repeats") == 2 and profile.get("warmup") == 0
            and profile.get("samples") == 5,
            "profile counts differ")
    require(profile.get("jobs") == [
        "xls-owned", "cfb-tiny", "cfb-many-small", "cfb-few-large"
    ], "profile jobs differ")
    require(profile.get("xls_owner") == XLS_OWNER, "XLS owner differs")
    require(profile.get("cfb_owner") == CFB_OWNER, "CFB owner differs")
    require(profile.get("instruction_flags") == list(PROFILE_FLAGS),
            "Callgrind instruction flags differ")

    allocation = plan.get("allocation")
    require(isinstance(allocation, dict), "allocation plan is missing")
    require(allocation.get("repeats") == 2 and allocation.get("warmup") == 3
            and allocation.get("samples") == 30,
            "allocation counts differ")
    native = plan.get("native")
    require(isinstance(native, dict), "native plan is missing")
    require(native.get("repeats") == 2 and native.get("warmup") == 20
            and native.get("samples") == 1000,
            "native counts differ")
    return plan


def matches(actual: str, expected: str) -> bool:
    """Match demangled names while tolerating generic/hash formatting."""

    if actual == expected:
        return True
    if actual.endswith("::" + expected.rsplit("::", 1)[-1]):
        return True
    return expected in actual


def marker_matches(name: str, markers: tuple[str, ...]) -> bool:
    return all(marker in name for marker in markers)


def option(command: list[str], name: str) -> str:
    values: list[str] = []
    for index, item in enumerate(command):
        if item == name:
            require(index + 1 < len(command), f"{name} has no value")
            values.append(command[index + 1])
        elif item.startswith(name + "="):
            values.append(item.split("=", 1)[1])
    require(len(values) == 1, f"{name} is missing or repeated")
    return values[0]


def profile_jobs(plan: dict[str, Any]) -> list[dict[str, Any]]:
    jobs: list[dict[str, Any]] = []
    profile = plan["profile"]
    for repeat in range(1, profile["repeats"] + 1):
        jobs.append({
            "name": f"profile-r{repeat}-xls-owned",
            "repeat": repeat,
            "group": "xls-owned",
            "kind": "xls",
            "shape": None,
            "owner": profile["xls_owner"],
            "runner": XLS_RUNNER,
            "setup_callers": [],
            "selection": {"cases": [XLS_CASE]},
            "samples": profile["samples"],
            "warmup": profile["warmup"],
        })
        for shape in CFB_SHAPES:
            jobs.append({
                "name": f"profile-r{repeat}-cfb-{shape}",
                "repeat": repeat,
                "group": f"cfb-{shape}",
                "kind": "cfb",
                "shape": shape,
                "owner": profile["cfb_owner"],
                "runner": CFB_RUNNER,
                "setup_callers": list(CFB_SETUP_CALLERS),
                "selection": {
                    "cases": [CFB_CASE],
                    "shapes": [shape],
                    "payload": "incompressible",
                },
                "samples": profile["samples"],
                "warmup": profile["warmup"],
            })
    return jobs


def validate_receipt(stage: str, job: dict[str, Any], plan: dict[str, Any]) -> dict[str, Any]:
    folder = HERE / stage
    name = job["name"]
    receipt_path = folder / f"{name}.receipt.json"
    receipt = read_json(receipt_path)
    require(isinstance(receipt, dict), f"{relative(receipt_path)} is not an object")
    require(receipt.get("exit_code") == 0, f"{name} exited unsuccessfully")
    # The matched profile lane follows the frozen ABBA order.  Baseline-r2
    # remains in the baseline output folder and uses the retained baseline
    # binary, while its live checkout/execution manifest is candidate.  Keep
    # the folder/source bindings separate so the folder name cannot silently
    # change the execution identity.
    expected_execution_stage = (
        "candidate" if stage == "baseline" and job["repeat"] == 2 else stage
    )
    require(receipt.get("execution_stage") == expected_execution_stage,
            f"{name} execution stage differs")
    expected_execution_manifest = HERE / expected_execution_stage / "source-manifest.json"
    require(receipt.get("execution_manifest_sha256") == sha256(expected_execution_manifest),
            f"{name} execution manifest binding differs")
    require(receipt.get("plan_sha256") == sha256(PLAN_PATH),
            f"{name} plan binding differs")
    require(receipt.get("script_sha256") == sha256(RUN_PATH),
            f"{name} driver binding differs")
    manifest_path = folder / "source-manifest.json"
    require(receipt.get("source_manifest_sha256") == sha256(manifest_path),
            f"{name} source manifest binding differs")

    command = receipt.get("command")
    require(isinstance(command, list) and all(isinstance(x, str) for x in command),
            f"{name} command is not argv")
    require(command[:7] == [
        "taskset", "-c", str(plan["cpu"]), "valgrind", "--vgdb=no",
        "--tool=callgrind", "--collect-atstart=no",
    ], f"{name} Callgrind prefix differs")
    require(command.count("--vgdb=no") == 1, f"{name} repeats --vgdb=no")
    for flag in PROFILE_FLAGS:
        require(flag in command and command.count(flag) == 1,
                f"{name} profile flag differs: {flag}")
    owner = job["owner"]
    for prefix in ("--toggle-collect=", "--zero-before=", "--dump-after="):
        require(prefix + owner in command and command.count(prefix + owner) == 1,
                f"{name} owner control differs: {prefix}")
    require(option(command, "--callgrind-out-file") ==
            str(folder / f"{name}.callgrind"),
            f"{name} Callgrind output path differs")
    require(option(command, "--case") == ",".join(job["selection"]["cases"]),
            f"{name} case differs")
    require(option(command, "--warmup") == str(job["warmup"]),
            f"{name} warmup differs")
    require(option(command, "--samples") == str(job["samples"]),
            f"{name} samples differs")
    require(option(command, "--json") == str(folder / f"{name}.json"),
            f"{name} JSON output differs")
    require(option(command, "--corpus-manifest") ==
            str(folder / f"{name}.catalog.json"),
            f"{name} corpus output differs")
    if job["shape"] is None:
        require("--shape" not in command and "--payload" not in command,
                f"{name} XLS command has CFB options")
    else:
        require(option(command, "--shape") == job["shape"],
                f"{name} shape differs")
        require(option(command, "--payload") == "incompressible",
                f"{name} payload differs")

    artifacts = receipt.get("artifacts")
    require(isinstance(artifacts, dict), f"{name} artifacts are missing")
    required = [
        f"{name}.host.json", f"{name}.json", f"{name}.catalog.json",
        f"{name}.stdout", f"{name}.stderr", f"{name}.callgrind",
    ]
    expected_numbered = 5 if job["kind"] == "xls" else 6
    required.extend(f"{name}.callgrind.{number}"
                    for number in range(1, expected_numbered + 1))
    for artifact in required:
        digest = artifacts.get(artifact)
        path = folder / artifact
        require(isinstance(digest, str) and len(digest) == 64,
                f"{name} artifact hash is missing: {artifact}")
        require(path.is_file() and not path.is_symlink(),
                f"{name} artifact is missing: {artifact}")
        require(sha256(path) == digest,
                f"{name} artifact hash differs: {artifact}")
    return {
        "path": relative(receipt_path),
        "sha256": sha256(receipt_path),
        "execution_stage": receipt["execution_stage"],
        "binary_sha256": receipt.get("binary_sha256"),
        "source_manifest_sha256": receipt["source_manifest_sha256"],
        "command": command,
        "artifacts": {key: artifacts[key] for key in sorted(artifacts)},
    }


def parse_part_and_trigger(text: str, label: str) -> tuple[int, str]:
    parts = [int(match.group(1)) for match in
             (PART_RE.match(line.strip()) for line in text.splitlines()) if match]
    triggers = [match.group(1) for match in
                (TRIGGER_RE.match(line.strip()) for line in text.splitlines()) if match]
    require(len(parts) == 1, f"{label}: expected one part record, got {parts}")
    require(len(triggers) == 1, f"{label}: expected one trigger, got {triggers}")
    return parts[0], triggers[0]


def positive_path(parsed: dict[str, Any], target_id: int,
                  ancestor: str) -> dict[str, Any] | None:
    """Find a bounded positive-cost path from an ancestor to target caller."""

    functions = parsed["functions"]
    starts = sorted(
        function_id for function_id, function in functions.items()
        if matches(function.get("name", ""), ancestor)
    )
    if not starts:
        return None
    queue: list[tuple[int, list[int], list[dict[str, Any]]]] = [
        (function_id, [function_id], []) for function_id in starts
    ]
    visited = set(starts)
    while queue:
        current_id, ids, edges = queue.pop(0)
        if current_id == target_id:
            return {
                "ancestor": ancestor,
                "function_ids": ids,
                "functions": [functions.get(function_id, {}).get("name", "")
                              for function_id in ids],
                "edges": edges,
                "depth": len(edges),
                "max_depth": MAX_ANCESTRY_DEPTH,
            }
        if len(edges) >= MAX_ANCESTRY_DEPTH:
            continue
        current = functions.get(current_id)
        if current is None:
            continue
        candidates = [
            edge for edge in current.get("edges", [])
            if edge.get("calls") is not None
            and edge["calls"] >= 0
            and edge.get("inclusive_ir", 0) > 0
        ]
        candidates.sort(key=lambda edge: (
            functions.get(edge["callee_id"], {}).get("name", ""),
            edge["callee_id"], edge.get("line") or 0,
        ))
        for edge in candidates:
            callee_id = edge["callee_id"]
            if callee_id in visited:
                continue
            visited.add(callee_id)
            queue.append((
                callee_id,
                [*ids, callee_id],
                [*edges, {
                    "caller_id": current_id,
                    "caller": functions.get(current_id, {}).get("name", ""),
                    "callee_id": callee_id,
                    "callee": functions.get(callee_id, {}).get("name", ""),
                    "calls": edge.get("calls"),
                    "inclusive_ir": edge.get("inclusive_ir", 0),
                    "line": edge.get("line"),
                }],
            ))
    return None


def context_path(parsed: dict[str, Any], function_id: int,
                 ancestor: str) -> dict[str, Any]:
    function = parsed["functions"].get(function_id, {})
    return {
        "ancestor": ancestor,
        "function_ids": [function_id],
        "functions": [function.get("name", "")],
        "edges": [],
        "depth": 0,
        "max_depth": MAX_ANCESTRY_DEPTH,
        "direct_owner_caller": True,
    }


def classify_owner(parsed: dict[str, Any], job: dict[str, Any], number: int,
                   label: str) -> dict[str, Any]:
    edge = parsed["owner_incoming"]
    caller = edge["caller"]
    caller_id = edge["caller_id"]
    direct_runner = matches(caller, job["runner"])
    direct_setup = [
        expected for expected in job["setup_callers"] if matches(caller, expected)
    ]
    require(not (direct_runner and direct_setup),
            f"{label}: owner caller matches both timed and setup roles")

    if direct_runner:
        role = "timed"
        runner_ancestry = context_path(parsed, caller_id, job["runner"])
        setup_ancestry = None
    elif direct_setup:
        role = "setup"
        runner_ancestry = None
        setup_ancestry = context_path(parsed, caller_id, direct_setup[0])
    else:
        runner_ancestry = positive_path(parsed, caller_id, job["runner"])
        setup_candidates = [
            positive_path(parsed, caller_id, expected)
            for expected in job["setup_callers"]
        ]
        setup_ancestry = next((candidate for candidate in setup_candidates
                               if candidate is not None), None)
        require(not (runner_ancestry is not None and setup_ancestry is not None),
                f"{label}: owner caller is reachable from timed and setup paths")
        if runner_ancestry is not None:
            role = "timed"
        elif setup_ancestry is not None:
            role = "setup"
        else:
            # This explicit wrapper branch gives a useful error for codegen
            # changes while retaining the wrapper as an allowed XLS boundary.
            if job["kind"] == "xls" and matches(caller, XLS_OWNER_WRAPPER):
                raise EvidenceError(
                    f"{label}: accepted XLS wrapper has no positive runner ancestry"
                )
            raise EvidenceError(
                f"{label}: owner caller {caller!r} has no allowed positive ancestry"
            )

    require(edge["calls"] == 1,
            f"{label}: selected owner edge has {edge['calls']} calls, expected 1")
    if role == "setup":
        require(job["kind"] == "cfb",
                f"{label}: XLS dump was classified as setup")
        require(number == 1,
                f"{label}: setup role appeared at numbered part {number}")
    else:
        require(job["kind"] == "xls" or number != 1,
                f"{label}: first CFB part was classified as timed")
    return {
        "role": role,
        "owner_caller": caller,
        "owner_caller_id": caller_id,
        "owner_calls": edge["calls"],
        "owner_incoming": edge,
        "runner_ancestry": runner_ancestry,
        "setup_ancestry": setup_ancestry,
        "validation": {
            "positive_owner_edge": True,
            "positive_ancestry_classified": True,
            "single_owner_call": True,
            "role_not_inferred_from_ordinal": True,
        },
    }


def target_ids(parsed: dict[str, Any], markers: tuple[str, ...]) -> list[int]:
    return sorted(
        function_id for function_id, function in parsed["functions"].items()
        if marker_matches(function.get("name", ""), markers)
    )


def target_attribution(parsed: dict[str, Any], key: str) -> dict[str, Any]:
    markers = TARGET_MARKERS[key]
    ids = target_ids(parsed, markers)
    functions = parsed["functions"]
    incoming: list[dict[str, Any]] = []
    selected = set(ids)
    for parent_id, parent in functions.items():
        for edge in parent.get("edges", []):
            if (edge.get("callee_id") in selected
                    and edge.get("calls") is not None
                    and edge["calls"] > 0
                    and edge.get("inclusive_ir", 0) > 0):
                incoming.append({
                    "caller_id": parent_id,
                    "caller": parent.get("name", ""),
                    "callee_id": edge["callee_id"],
                    "callee": functions.get(edge["callee_id"], {}).get("name", ""),
                    "calls": edge["calls"],
                    "inclusive_ir": edge["inclusive_ir"],
                    "line": edge.get("line"),
                })
    incoming.sort(key=lambda edge: (
        edge["caller"], edge["callee"], edge["callee_id"], edge["line"] or 0,
    ))
    self_ir = sum(functions[function_id].get("self_ir", 0) for function_id in ids)
    direct_ir = sum(
        edge.get("inclusive_ir", 0)
        for function_id in ids
        for edge in functions[function_id].get("edges", [])
    )
    if not ids:
        state = "absent"
        absence_reason = "function_absent"
    elif not incoming:
        state = "inlined_or_unreferenced"
        absence_reason = "no_positive_incoming_edge"
    else:
        state = "out_of_line"
        absence_reason = None
    return {
        "target": key,
        "markers": list(markers),
        "function_ids": ids,
        "state": state,
        "absence_reason": absence_reason,
        "self_ir": self_ir,
        "direct_ir": direct_ir,
        "inclusive_ir": sum(edge["inclusive_ir"] for edge in incoming),
        "calls": sum(edge["calls"] for edge in incoming),
        "incoming_edge_count": len(incoming),
        "incoming_edges": incoming,
        "interpretation": (
            "Named emitted function with positive incoming edge."
            if state == "out_of_line" else
            "Code shape is absent or has no positive edge; inspect owner disassembly "
            "before inferring moved or eliminated work."
        ),
    }


def direct_edges(parsed: dict[str, Any], caller_markers: tuple[str, ...],
                 callee_markers: tuple[str, ...]) -> dict[str, Any]:
    functions = parsed["functions"]
    callee_ids = set(target_ids(parsed, callee_markers))
    rows: list[dict[str, Any]] = []
    for caller_id, caller in functions.items():
        if not marker_matches(caller.get("name", ""), caller_markers):
            continue
        for edge in caller.get("edges", []):
            if (edge.get("callee_id") in callee_ids
                    and edge.get("calls") is not None
                    and edge["calls"] > 0
                    and edge.get("inclusive_ir", 0) > 0):
                rows.append({
                    "caller_id": caller_id,
                    "caller": caller.get("name", ""),
                    "callee_id": edge["callee_id"],
                    "callee": functions.get(edge["callee_id"], {}).get("name", ""),
                    "calls": edge["calls"],
                    "inclusive_ir": edge["inclusive_ir"],
                    "line": edge.get("line"),
                })
    rows.sort(key=lambda row: (
        row["caller"], row["callee"], row["callee_id"], row["line"] or 0,
    ))
    return {
        "caller_markers": list(caller_markers),
        "callee_markers": list(callee_markers),
        "edges": rows,
        "positive_edge_count": len(rows),
        "calls": sum(row["calls"] for row in rows),
        "inclusive_ir": sum(row["inclusive_ir"] for row in rows),
    }


def numbered_paths(folder: Path, name: str) -> list[tuple[int, Path]]:
    prefix = f"{name}.callgrind."
    found: list[tuple[int, Path]] = []
    for path in folder.glob(prefix + "*"):
        suffix = path.name[len(prefix):]
        if suffix.isdigit() and path.is_file() and not path.is_symlink():
            found.append((int(suffix), path))
    found.sort()
    require(found, f"{relative(folder / name)}: no numbered Callgrind dumps")
    require([number for number, _ in found] == list(range(1, len(found) + 1)),
            f"{name}: numbered dumps are not contiguous")
    return found


def final_dump(path: Path, expected_part: int, owner: str) -> dict[str, Any]:
    text = read_text(path)
    label = relative(path)
    part, trigger = parse_part_and_trigger(text, label)
    require(part == expected_part, f"{label}: final part is {part}, expected {expected_part}")
    require(trigger == "Program termination", f"{label}: final trigger differs")
    events = [line.strip().split(":", 1)[1].split()
              for line in text.splitlines() if line.strip().startswith("events:")]
    positions = [line.strip().split(":", 1)[1].split()
                 for line in text.splitlines() if line.strip().startswith("positions:")]
    require(events == [["Ir"]], f"{label}: final event set differs")
    require(positions == [["instr"]], f"{label}: final positions differ")
    summaries = []
    for line in text.splitlines():
        match = SUMMARY_RE.match(line.strip())
        if match:
            token = match.group(1).split()[0]
            try:
                summaries.append(int(token.replace(",", "")))
            except ValueError as error:
                raise EvidenceError(f"{label}: malformed final summary") from error
    require(summaries and summaries[0] == 0,
            f"{label}: final summary Ir is not zero")
    return {
        "file": label,
        "sha256": sha256(path),
        "part": part,
        "trigger": trigger,
        "summary_ir": summaries[0],
        "owner": owner,
        "validation": {"program_termination_zero_ir": True},
    }


def parse_dump(stage: str, job: dict[str, Any], number: int,
               path: Path) -> dict[str, Any]:
    label = relative(path)
    text = read_text(path)
    part, trigger = parse_part_and_trigger(text, label)
    require(part == number, f"{label}: part {part} does not match suffix {number}")
    require(trigger == f"--dump-after={job['owner']}",
            f"{label}: trigger does not bind selected owner")
    events = [line.strip().split(":", 1)[1].split()
              for line in text.splitlines() if line.strip().startswith("events:")]
    positions = [line.strip().split(":", 1)[1].split()
                 for line in text.splitlines() if line.strip().startswith("positions:")]
    require(events == [["Ir"]], f"{label}: event set differs")
    require(positions == [["instr"]], f"{label}: positions differ")
    parsed = RAW.parse_raw_profile(path.resolve(), job["owner"])
    classification = classify_owner(parsed, job, number, label)
    attributions = {
        key: target_attribution(parsed, key) for key in TARGET_MARKERS
    }
    decoder_edge = direct_edges(
        parsed, TARGET_MARKERS["parse_directory_entry"],
        TARGET_MARKERS["decode_utf16le"],
    )
    scalar_edge = direct_edges(
        parsed, TARGET_MARKERS["parse_directory_entry"],
        TARGET_MARKERS["format_clsid"],
    )
    expected = BASELINE_DECODER_CALLS[job["group"]]
    observed = decoder_edge["calls"]
    return {
        "number": number,
        "role": classification["role"],
        "owner_caller": classification["owner_caller"],
        "owner_calls": classification["owner_calls"],
        "owner_incoming": classification["owner_incoming"],
        "owner_self_ir": parsed["owner_self_ir"],
        "owner_direct_ir": parsed["owner_direct_ir"],
        "summary_ir": parsed["summary_ir"],
        "runner_ancestry": classification["runner_ancestry"],
        "setup_ancestry": classification["setup_ancestry"],
        "file": parsed["file"],
        "sha256": parsed["sha256"],
        "trigger": trigger,
        "events": events[0],
        "positions": positions[0],
        "attributions": attributions,
        "direct_edges": {
            "parse_directory_entry_to_decode_utf16le": decoder_edge,
            "parse_directory_entry_to_format_clsid": scalar_edge,
        },
        "decoder_prediction": {
            "expected_baseline_calls": expected,
            "observed_positive_direct_calls": observed,
            "observed_positive_edge_count": decoder_edge["positive_edge_count"],
            "baseline_matches_corpus_prediction": (
                stage != "baseline" or observed == expected
            ),
            "candidate_exactly_one_standard_root_call": (
                stage != "candidate" or observed == 1
            ),
            "interpretation": (
                "The direct parse_directory_entry incoming edge is the decoder "
                "call count; decode_utf16le outgoing try_reserve edges are not calls."
            ),
        },
        "scalar_control": {
            "format_clsid_direct_calls": scalar_edge["calls"],
            "format_clsid_direct_ir": scalar_edge["inclusive_ir"],
            "format_clsid_attribution_state": attributions["format_clsid"]["state"],
            "interpretation": (
                "Scalar/CLSID work is retained as a control; absent or inlined "
                "symbols need owner disassembly before any removal claim."
            ),
        },
        "validation": {
            **classification["validation"],
            "owner_edge_matches_summary": (
                classification["owner_incoming"]["inclusive_ir"] == parsed["summary_ir"]
            ),
            "owner_self_plus_direct_matches_summary": (
                parsed["owner_self_ir"] + parsed["owner_direct_ir"] == parsed["summary_ir"]
            ),
            "per_dump_decoder_and_scalar_evidence": True,
            "absence_is_explicit_not_zero_work": True,
        },
    }


def aggregate_rows(profiles: list[dict[str, Any]], role: str = "timed") -> dict[str, Any]:
    totals: dict[str, dict[str, int]] = {}
    decoder_by_job: dict[str, dict[str, Any]] = {}
    for profile in profiles:
        decoder_by_job[profile["name"]] = {
            "repeat": profile["repeat"],
            "group": profile["group"],
            "shape": profile["shape"],
            "dumps": [],
        }
        for dump in profile["raw_dumps"]:
            if dump["role"] != role:
                continue
            for key, attribution in dump["attributions"].items():
                row = totals.setdefault(key, {
                    "self_ir": 0,
                    "direct_ir": 0,
                    "inclusive_ir": 0,
                    "calls": 0,
                    "incoming_edge_count": 0,
                    "dump_count": 0,
                })
                for field in ("self_ir", "direct_ir", "inclusive_ir", "calls",
                              "incoming_edge_count"):
                    row[field] += int(attribution[field])
                row["dump_count"] += 1
            prediction = dump["decoder_prediction"]
            decoder_by_job[profile["name"]]["dumps"].append({
                "number": dump["number"],
                "observed_positive_direct_calls": prediction["observed_positive_direct_calls"],
                "expected_baseline_calls": prediction["expected_baseline_calls"],
                "parse_to_decode_ir": dump["direct_edges"][
                    "parse_directory_entry_to_decode_utf16le"
                ]["inclusive_ir"],
                "parse_to_format_clsid_calls": dump["direct_edges"][
                    "parse_directory_entry_to_format_clsid"
                ]["calls"],
            })
    return {
        "target_totals": totals,
        "decoder_by_job": decoder_by_job,
    }


def analyze_stage(stage: str, plan: dict[str, Any]) -> dict[str, Any]:
    require(stage in STAGES, f"unsupported stage {stage}")
    folder = HERE / stage
    profiles: list[dict[str, Any]] = []
    for job in profile_jobs(plan):
        receipt = validate_receipt(stage, job, plan)
        numbered = numbered_paths(folder, job["name"])
        expected_numbered = 5 if job["kind"] == "xls" else 6
        require(len(numbered) == expected_numbered,
                f"{job['name']}: numbered dump count differs")
        dumps = [parse_dump(stage, job, number, path)
                 for number, path in numbered]
        setup = [dump for dump in dumps if dump["role"] == "setup"]
        timed = [dump for dump in dumps if dump["role"] == "timed"]
        expected_setup = 0 if job["kind"] == "xls" else 1
        require(len(setup) == expected_setup,
                f"{job['name']}: setup count differs")
        require(len(timed) == plan["profile"]["samples"],
                f"{job['name']}: timed count differs")
        final_path = folder / f"{job['name']}.callgrind"
        final = final_dump(final_path, expected_numbered + 1, job["owner"])
        profiles.append({
            "name": job["name"],
            "repeat": job["repeat"],
            "group": job["group"],
            "kind": job["kind"],
            "shape": job["shape"],
            "owner": job["owner"],
            "runner": job["runner"],
            "receipt": receipt,
            "raw_dumps": dumps,
            "setup_dumps": setup,
            "timed_dumps": timed,
            "final_process_dump": final,
            "validation": {
                "actual_positive_ancestry_classifies_roles": True,
                "setup_excluded_from_timed_aggregates": True,
                "all_timed_owner_calls_one": all(
                    dump["owner_calls"] == 1 for dump in timed
                ),
                "all_owner_equations_match": all(
                    dump["validation"]["owner_edge_matches_summary"]
                    and dump["validation"]["owner_self_plus_direct_matches_summary"]
                    for dump in dumps
                ),
                "termination_dump_zero_ir": True,
            },
        })

    timed_count = sum(len(profile["timed_dumps"]) for profile in profiles)
    setup_count = sum(len(profile["setup_dumps"]) for profile in profiles)
    aggregate = aggregate_rows(profiles)
    candidate_exact = all(
        dump["decoder_prediction"]["candidate_exactly_one_standard_root_call"]
        for profile in profiles if stage == "candidate"
        for dump in profile["timed_dumps"]
    ) if stage == "candidate" else None
    baseline_predicted = all(
        dump["decoder_prediction"]["baseline_matches_corpus_prediction"]
        for profile in profiles if stage == "baseline"
        for dump in profile["timed_dumps"]
    ) if stage == "baseline" else None
    return {
        "schema": "ole2_name_handoff_0554_profile_analysis_v1",
        "status": "pass",
        "stage": stage,
        "scope": plan["scope"],
        "performance_claim": "diagnostic-only",
        "plan": relative(PLAN_PATH),
        "plan_sha256": sha256(PLAN_PATH),
        "profile_count": len(profiles),
        "timed_constructor_dump_count": timed_count,
        "setup_dump_count": setup_count,
        "profiles": profiles,
        "aggregates": aggregate,
        "mechanism_gate": {
            "all_timed_owner_edges_classified": all(
                profile["validation"]["actual_positive_ancestry_classifies_roles"]
                for profile in profiles
            ),
            "baseline_direct_decoder_calls_match_corpus_prediction": baseline_predicted,
            "candidate_standard_root_direct_decoder_calls_exactly_one": candidate_exact,
            "scalar_and_clsid_attribution_retained_separately": True,
            "missing_or_inlined_symbols_require_disassembly": True,
        },
        "helpers": {
            "change-0536/analyze_profiles.py": sha256(RAW_HELPER_PATH),
            "run.py": sha256(RUN_PATH),
        },
        "validation": {
            "plan_and_receipts_bound": True,
            "raw_dumps_hashed_and_parsed": True,
            "positive_owner_edge_and_ancestry_used": True,
            "setup_and_termination_dumps_retained": True,
            "decoder_incoming_edge_direction_correct": True,
            "scalar_fields_separate_from_name_decoder": True,
            "no_symbol_absence_zero_work_claim": True,
            "no_native_or_adoption_claim": True,
        },
        "limitations": [
            "Callgrind Ir and collection-off call metadata are mechanism diagnostics, not native latency or operation-local allocation counts.",
            "An absent or inlined function is not evidence that its work disappeared; owner disassembly and instruction-position mapping are required.",
        ],
    }


def metric(baseline: int, candidate: int) -> dict[str, int | float | None]:
    delta = candidate - baseline
    return {
        "baseline": baseline,
        "candidate": candidate,
        "delta": delta,
        "delta_percent": None if baseline == 0 else (100.0 * delta / baseline),
    }


def profile_key(profile: dict[str, Any]) -> tuple[int, str, str | None]:
    return profile["repeat"], profile["group"], profile["shape"]


def compare_stages(baseline: dict[str, Any], candidate: dict[str, Any]) -> dict[str, Any]:
    b_profiles = {profile_key(profile): profile for profile in baseline["profiles"]}
    c_profiles = {profile_key(profile): profile for profile in candidate["profiles"]}
    require(set(b_profiles) == set(c_profiles), "stage profile matrices differ")
    rows: list[dict[str, Any]] = []
    target_keys = tuple(TARGET_MARKERS)
    for key in sorted(b_profiles):
        bp = b_profiles[key]
        cp = c_profiles[key]
        bt = aggregate_rows([bp])["target_totals"]
        ct = aggregate_rows([cp])["target_totals"]
        target_metrics: dict[str, Any] = {}
        for target in target_keys:
            b = bt.get(target, {field: 0 for field in
                                ("self_ir", "direct_ir", "inclusive_ir", "calls",
                                 "incoming_edge_count", "dump_count")})
            c = ct.get(target, {field: 0 for field in
                                ("self_ir", "direct_ir", "inclusive_ir", "calls",
                                 "incoming_edge_count", "dump_count")})
            target_metrics[target] = {
                field: metric(int(b[field]), int(c[field]))
                for field in ("self_ir", "direct_ir", "inclusive_ir", "calls",
                              "incoming_edge_count", "dump_count")
            }
        b_decoder = [
            dump["decoder_prediction"]["observed_positive_direct_calls"]
            for dump in bp["timed_dumps"]
        ]
        c_decoder = [
            dump["decoder_prediction"]["observed_positive_direct_calls"]
            for dump in cp["timed_dumps"]
        ]
        b_scalar = [
            dump["scalar_control"]["format_clsid_direct_calls"]
            for dump in bp["timed_dumps"]
        ]
        c_scalar = [
            dump["scalar_control"]["format_clsid_direct_calls"]
            for dump in cp["timed_dumps"]
        ]
        rows.append({
            "repeat": key[0],
            "group": key[1],
            "shape": key[2],
            "owner_incoming_ir": metric(
                sum(dump["summary_ir"] for dump in bp["timed_dumps"]),
                sum(dump["summary_ir"] for dump in cp["timed_dumps"]),
            ),
            "targets": target_metrics,
            "decoder_calls_per_timed_dump": {
                "baseline": b_decoder,
                "candidate": c_decoder,
                "candidate_all_exactly_one": all(value == 1 for value in c_decoder),
            },
            "format_clsid_calls_per_timed_dump": {
                "baseline": b_scalar,
                "candidate": c_scalar,
                "same_call_vector": b_scalar == c_scalar,
            },
        })

    many_small = [
        row for row in rows
        if row["group"] == "cfb-many-small"
    ]
    return {
        "rows": rows,
        "mechanism_gate": {
            "candidate_standard_root_decoder_exactly_one_all_dumps": all(
                row["decoder_calls_per_timed_dump"]["candidate_all_exactly_one"]
                for row in rows
            ),
            "baseline_decoder_predictions_match": all(
                all(value == BASELINE_DECODER_CALLS[row["group"]]
                    for value in row["decoder_calls_per_timed_dump"]["baseline"])
                for row in rows
            ),
            "many_small_owner_ir_decreases_each_repeat": all(
                row["owner_incoming_ir"]["delta"] < 0 for row in many_small
            ),
            "format_clsid_call_vector_preserved": all(
                row["format_clsid_calls_per_timed_dump"]["same_call_vector"]
                for row in rows
            ),
            "name_decoder_and_scalar_rows_are_separate": True,
        },
        "interpretation": (
            "Negative Ir deltas are mechanism diagnostics only.  A candidate "
            "decoder call vector of one is the predicted standard-root shape; "
            "instruction mapping is still required to establish removed work."
        ),
    }


def compare(plan: dict[str, Any]) -> dict[str, Any]:
    baseline = analyze_stage("baseline", plan)
    candidate = analyze_stage("candidate", plan)
    return {
        "schema": "ole2_name_handoff_0554_profile_comparison_v1",
        "status": "pass",
        "scope": plan["scope"],
        "performance_claim": "diagnostic-only",
        "plan": relative(PLAN_PATH),
        "plan_sha256": sha256(PLAN_PATH),
        "stage_selection": ["baseline", "candidate"],
        "stages": {"baseline": baseline, "candidate": candidate},
        "comparison": compare_stages(baseline, candidate),
        "validation": {
            "both_stages_valid": True,
            "matched_owner_edges_compared": True,
            "per_dump_decoder_prediction_compared": True,
            "scalar_clsid_control_compared": True,
            "no_native_or_adoption_claim": True,
        },
        "limitations": [
            "Callgrind Ir is a mechanism diagnostic and does not establish native latency, RSS, allocation, hardware, or scaling behavior.",
            "A missing or inlined named helper needs static owner disassembly and instruction-position evidence before any work-removal interpretation.",
        ],
    }


def write_report(document: dict[str, Any], output: Path) -> None:
    data = (json.dumps(document, indent=2, sort_keys=True) + "\n").encode("utf-8")
    if output.exists():
        require(output.is_file() and not output.is_symlink(),
                f"refusing non-regular output {output}")
        require(output.read_bytes() == data,
                f"refusing to overwrite non-identical output {output}")
        return
    require(not (HERE / "SHA256SUMS").exists() or
            not output.resolve().is_relative_to(HERE.resolve()),
            f"refusing new report in sealed bundle {output}")
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_bytes(data)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("output", nargs="?", type=Path)
    parser.add_argument("--output", dest="output_option", type=Path)
    parser.add_argument("--stage", choices=STAGES, default="baseline")
    parser.add_argument("--compare", action="store_true")
    args = parser.parse_args(argv)
    require_not_both = not (args.output is not None and args.output_option is not None)
    if not require_not_both:
        parser.error("provide output positionally or with --output, not both")
    output = (args.output_option or args.output or
              (HERE / "profile-comparison.json" if args.compare else
               HERE / args.stage / "profile-analysis.json"))
    try:
        plan = plan_data()
        document = compare(plan) if args.compare else analyze_stage(args.stage, plan)
        write_report(document, output)
    except (EvidenceError, OSError, TypeError, ValueError, KeyError) as error:
        print(f"analyze_profiles.py: error: {error}", file=sys.stderr)
        return 2
    print(output)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
