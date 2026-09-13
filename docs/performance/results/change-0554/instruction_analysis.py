#!/usr/bin/env python3
"""Map 0554 Callgrind self-instruction costs to bounded static symbols.

This is a mechanism report for the matched OLE2 name-handoff experiment.  It
uses the immutable 0547 parser for ``positions: instr`` Callgrind records and
for the exact ``objdump`` output retained by ``inspect_assembly.py``.  The
0554 plan, receipt, role, and stage checks live here; no old campaign verifier
or target path is consulted.

Only exclusive function self instruction Ir is added to a target total.
Inclusive cfn edge Ir and call counts are emitted as separate evidence.  A
missing or inlined symbol is recorded as such and never converted to zero
work.  The mapper is deliberately bounded to the exact symbols selected by
the stage assembly index; an owner/caller disassembly is needed for code that
has been inlined into a selected parent.
"""

from __future__ import annotations

import argparse
import collections
import hashlib
import importlib.util
import json
from pathlib import Path
import sys
from typing import Any, Iterable

sys.dont_write_bytecode = True


HERE = Path(__file__).resolve().parent
PLAN_PATH = HERE / "plan.json"
RUN_PATH = HERE / "run.py"
INSPECT_PATH = HERE / "inspect_assembly.py"
SCRIPT_PATH = HERE / "instruction_analysis.py"
RAW_PARSER_PATH = HERE.parent / "change-0547" / "instruction_analysis.py"
STAGES = ("baseline", "candidate")

XLS_OWNER = (
    "litchi_xls::workbook::source::SourceBackedWorkbook::from_read_at_with_limits"
)
CFB_OWNER = "litchi_cfb::file::OleFile<R>::open"
XLS_RUNNER = "litchi_perf_baseline::run_xls_owned_source_case"
CFB_RUNNER = "litchi_perf_baseline::run_cfb_open"
CFB_SETUP_CALLERS = (
    "litchi_perf_baseline::build_cfb_corpus",
    "litchi_perf_baseline::run::{{closure}}",
)
XLS_OWNER_WRAPPER = "litchi_xls::workbook::source::SourceBackedWorkbook::from_read_at"
CFB_SHAPES = ("tiny", "many-small", "few-large")
XLS_CASE = "xls_owned_source_open_one_cell"
CFB_CASE = "cfb_open"
MAX_ANCESTRY_DEPTH = 8

# These are the only named functions for which this report makes an
# instruction-position mapping.  The index may also contain validator/tree
# symbols; their presence is retained in the static inventory, while dynamic
# totals for those optional/inlined helpers remain in profile-analysis.json.
TARGET_MARKERS: dict[str, tuple[str, ...]] = {
    "parse_directory_entry": ("::parse_directory_entry",),
    "decode_utf16le": ("::decode_utf16le",),
    "format_clsid": ("::format_clsid",),
}

ALL_MARKERS: dict[str, tuple[str, ...]] = {
    **TARGET_MARKERS,
    "validated_directory_entries": ("::validated_directory_entries",),
    "parse_validated_directory_entry": ("::parse_validated_directory_entry",),
    "build_storage_tree_iterative": ("::build_storage_tree_iterative",),
    "directory_name_data": ("directory_name::directory_name_data",),
    "name_data_drop": ("DirectoryNameData", "drop_in_place"),
}

# ``positions: instr`` function names are demangled, while the assembly
# index intentionally retains the exact mangled nm symbol.  Keep separate
# fragments for that boundary; using the demangled ``::`` markers against an
# nm symbol would falsely classify every mapped instruction as outside its
# selected symbol.
STATIC_MARKERS: dict[str, tuple[str, ...]] = {
    "parse_directory_entry": ("parse_directory_entry",),
    "decode_utf16le": ("decode_utf16le",),
    "format_clsid": ("format_clsid",),
    "validated_directory_entries": ("validated_directory_entries",),
    "parse_validated_directory_entry": ("parse_validated_directory_entry",),
    "build_storage_tree_iterative": ("build_storage_tree_iterative",),
    "directory_name_data": ("directory_name_data",),
    "name_data_drop": ("DirectoryNameData", "drop_in_place"),
}

PROFILE_FLAGS = (
    "--dump-instr=yes",
    "--dump-line=no",
    "--compress-pos=no",
    "--collect-jumps=yes",
)


class EvidenceError(ValueError):
    """A missing, malformed, or contradictory evidence artifact."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise EvidenceError(message)


def sha256(path: Path) -> str:
    try:
        digest = hashlib.sha256()
        with path.open("rb") as stream:
            for block in iter(lambda: stream.read(1024 * 1024), b""):
                digest.update(block)
        return digest.hexdigest()
    except OSError as error:
        raise EvidenceError(f"cannot hash {path}: {error}") from error


def read_json(path: Path) -> Any:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise EvidenceError(f"cannot read JSON {path}: {error}") from error


def read_text(path: Path) -> str:
    try:
        return path.read_text(encoding="utf-8", errors="replace")
    except OSError as error:
        raise EvidenceError(f"cannot read {path}: {error}") from error


def relative(path: Path) -> str:
    try:
        return path.relative_to(HERE).as_posix()
    except ValueError as error:
        raise EvidenceError(f"path is outside evidence directory: {path}") from error


def load_immutable_parser() -> Any:
    require(RAW_PARSER_PATH.is_file(), f"missing immutable parser {RAW_PARSER_PATH}")
    spec = importlib.util.spec_from_file_location(
        "litchi_0554_immutable_instruction_parser", RAW_PARSER_PATH
    )
    require(spec is not None and spec.loader is not None,
            f"cannot load immutable parser {RAW_PARSER_PATH}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    # The helper only uses HERE in diagnostics.  It does not load a campaign
    # plan or verifier state, so this rebinding keeps its parsing routines
    # reusable without importing the old campaign boundary.
    module.HERE = HERE
    return module


RAW = load_immutable_parser()


def marker_matches(name: str, markers: Iterable[str]) -> bool:
    return all(marker in name for marker in markers)


def static_marker_matches(symbol: str, target: str) -> bool:
    # The short mangled fragment ``parse_directory_entry`` is also contained
    # in ``parse_validated_directory_entry``.  Keep those bounds disjoint if a
    # later build emits both helpers.
    if target == "parse_directory_entry" and "parse_validated_directory_entry" in symbol:
        return False
    return marker_matches(symbol, STATIC_MARKERS[target])


def name_matches(actual: str, expected: str) -> bool:
    return actual == expected or actual.endswith("::" + expected) or expected in actual


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


def load_plan() -> dict[str, Any]:
    plan = read_json(PLAN_PATH)
    require(isinstance(plan, dict), "plan is not an object")
    require(plan.get("schema") == "ole2_name_handoff_0554_plan_v1",
            "plan schema differs")
    require(plan.get("scope") ==
            "Matched private CFB validated-name ownership handoff N; NF field seed excluded",
            "plan scope differs")
    require(plan.get("cpu") == 2, "plan CPU differs")
    require(plan.get("candidate_files") == ["crates/litchi-cfb/src/file.rs"],
            "candidate scope differs")
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
    cfb = plan.get("groups", {}).get("cfb", {})
    require(cfb.get("cases") == [CFB_CASE], "CFB case matrix differs")
    require(tuple(cfb.get("shapes", ())) == CFB_SHAPES,
            "CFB shape matrix differs")
    require(cfb.get("payload") == "incompressible", "CFB payload differs")
    require(XLS_CASE in plan.get("groups", {}).get("xls", {}).get("cases", []),
            "XLS one-cell case is absent")
    return plan


def profile_jobs(plan: dict[str, Any]) -> list[dict[str, Any]]:
    profile = plan["profile"]
    jobs: list[dict[str, Any]] = []
    for repeat in range(1, int(profile["repeats"]) + 1):
        jobs.append({
            "name": f"profile-r{repeat}-xls-owned",
            "repeat": repeat,
            "group": "xls-owned",
            "kind": "xls",
            "shape": None,
            "owner": profile["xls_owner"],
            "runner": XLS_RUNNER,
            "setup_callers": [],
            "cases": [XLS_CASE],
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
                "cases": [CFB_CASE],
            })
    return jobs


def validate_receipt(stage: str, job: dict[str, Any], plan: dict[str, Any]) -> dict[str, Any]:
    folder = HERE / stage
    name = job["name"]
    receipt_path = folder / f"{name}.receipt.json"
    receipt = read_json(receipt_path)
    require(isinstance(receipt, dict), f"{relative(receipt_path)} is not an object")
    require(receipt.get("exit_code") == 0, f"{name} exited unsuccessfully")
    # Native/profile ABBA retains baseline-r2 in the baseline output folder
    # while binding its live source transition as execution-stage candidate.
    # The folder still owns the baseline binary and source-manifest hash; this
    # explicit exception preserves that distinction instead of accepting an
    # arbitrary folder-name/source mismatch.
    expected_execution_stage = (
        "candidate" if stage == "baseline" and job["repeat"] == 2 else stage
    )
    require(receipt.get("execution_stage") == expected_execution_stage,
            f"{name} execution stage differs")
    require(receipt.get("plan_sha256") == sha256(PLAN_PATH),
            f"{name} plan binding differs")
    require(receipt.get("script_sha256") == sha256(RUN_PATH),
            f"{name} driver binding differs")
    manifest_path = folder / "source-manifest.json"
    require(receipt.get("source_manifest_sha256") == sha256(manifest_path),
            f"{name} source manifest binding differs")
    command = receipt.get("command")
    require(isinstance(command, list) and all(isinstance(item, str) for item in command),
            f"{name} command is not argv")
    require(command[:7] == [
        "taskset", "-c", str(plan["cpu"]), "valgrind", "--vgdb=no",
        "--tool=callgrind", "--collect-atstart=no",
    ], f"{name} Callgrind prefix differs")
    for flag in PROFILE_FLAGS:
        require(command.count(flag) == 1, f"{name} profile flag differs: {flag}")
    for prefix in ("--toggle-collect=", "--zero-before=", "--dump-after="):
        require(command.count(prefix + job["owner"]) == 1,
                f"{name} owner control differs: {prefix}")
    folder_name = str(folder / f"{name}")
    require(option(command, "--callgrind-out-file") == folder_name + ".callgrind",
            f"{name} Callgrind output path differs")
    require(option(command, "--case") == ",".join(job["cases"]),
            f"{name} case differs")
    require(option(command, "--warmup") == str(plan["profile"]["warmup"]),
            f"{name} warmup differs")
    require(option(command, "--samples") == str(plan["profile"]["samples"]),
            f"{name} samples differs")
    require(option(command, "--json") == folder_name + ".json",
            f"{name} JSON output path differs")
    require(option(command, "--corpus-manifest") == folder_name + ".catalog.json",
            f"{name} corpus output path differs")
    if job["shape"] is None:
        require("--shape" not in command and "--payload" not in command,
                f"{name} has CFB-only options")
    else:
        require(option(command, "--shape") == job["shape"],
                f"{name} shape differs")
        require(option(command, "--payload") == "incompressible",
                f"{name} payload differs")

    artifacts = receipt.get("artifacts")
    require(isinstance(artifacts, dict), f"{name} artifacts are missing")
    numbered_count = 5 if job["kind"] == "xls" else 6
    required = [
        f"{name}.host.json", f"{name}.json", f"{name}.catalog.json",
        f"{name}.stdout", f"{name}.stderr", f"{name}.callgrind",
        *[f"{name}.callgrind.{number}" for number in range(1, numbered_count + 1)],
    ]
    for artifact in required:
        digest = artifacts.get(artifact)
        artifact_path = folder / artifact
        require(isinstance(digest, str) and len(digest) == 64,
                f"{name} artifact hash is missing: {artifact}")
        require(artifact_path.is_file() and not artifact_path.is_symlink(),
                f"{name} artifact is missing: {artifact}")
        require(sha256(artifact_path) == digest,
                f"{name} artifact hash differs: {artifact}")
    return {
        "file": relative(receipt_path),
        "sha256": sha256(receipt_path),
        "execution_stage": receipt["execution_stage"],
        "binary_sha256": receipt.get("binary_sha256"),
        "source_manifest_sha256": receipt["source_manifest_sha256"],
        "artifacts": {key: artifacts[key] for key in sorted(artifacts)},
    }


def parse_dump_role(parsed: dict[str, Any], job: dict[str, Any], number: int) -> dict[str, Any]:
    """Classify a dump from positive owner ancestry, independent of ordinal."""

    owner_matches = [
        (function_id, function) for function_id, function in parsed["functions"].items()
        if job["owner"] in function.get("name", "")
    ]
    require(len(owner_matches) == 1,
            f"{parsed['path']}: selected owner has {len(owner_matches)} function ids")
    owner_id, owner = owner_matches[0]
    incoming = []
    for parent in parsed["functions"].values():
        for edge in parent.get("edges", []):
            if (edge.get("callee_id") == owner_id
                    and edge.get("calls") is not None
                    and edge["calls"] >= 0
                    and edge.get("inclusive_ir", 0) > 0):
                incoming.append({
                    "caller_id": parent["id"],
                    "caller": parent.get("name", ""),
                    "callee_id": owner_id,
                    "callee": owner.get("name", ""),
                    "calls": edge["calls"],
                    "inclusive_ir": edge["inclusive_ir"],
                    "position": edge.get("position"),
                })
    require(len(incoming) == 1,
            f"{parsed['path']}: expected one positive owner edge, got {len(incoming)}")
    owner_edge = incoming[0]
    require(owner_edge["calls"] == 1,
            f"{parsed['path']}: owner edge calls {owner_edge['calls']}, expected 1")
    require(owner_edge["inclusive_ir"] == parsed["summary"],
            f"{parsed['path']}: owner edge does not match summary")
    direct_ir = sum(edge.get("inclusive_ir", 0) for edge in owner.get("edges", []))
    require(owner.get("self_ir", 0) + direct_ir == owner_edge["inclusive_ir"],
            f"{parsed['path']}: owner self plus direct Ir does not close")

    def positive_path(start_name: str, target_id: int) -> dict[str, Any] | None:
        starts = sorted(
            function_id for function_id, function in parsed["functions"].items()
            if name_matches(function.get("name", ""), start_name)
        )
        queue: collections.deque[tuple[int, list[int], list[dict[str, Any]]]] = (
            collections.deque((start, [start], []) for start in starts)
        )
        seen = set(starts)
        while queue:
            current, ids, edges = queue.popleft()
            if current == target_id:
                return {
                    "ancestor": start_name,
                    "function_ids": ids,
                    "functions": [parsed["functions"][ident]["name"] for ident in ids],
                    "edges": edges,
                    "depth": len(edges),
                    "max_depth": MAX_ANCESTRY_DEPTH,
                }
            if len(edges) >= MAX_ANCESTRY_DEPTH:
                continue
            function = parsed["functions"].get(current)
            if function is None:
                continue
            candidates = [
                edge for edge in function.get("edges", [])
                if edge.get("calls") is not None
                and edge["calls"] >= 0
                and edge.get("inclusive_ir", 0) > 0
            ]
            candidates.sort(key=lambda edge: (
                parsed["functions"].get(edge["callee_id"], {}).get("name", ""),
                edge["callee_id"],
            ))
            for edge in candidates:
                callee_id = edge["callee_id"]
                if callee_id in seen:
                    continue
                seen.add(callee_id)
                edge_record = {
                    "caller_id": current,
                    "caller": function.get("name", ""),
                    "callee_id": callee_id,
                    "callee": parsed["functions"].get(callee_id, {}).get("name", ""),
                    "calls": edge["calls"],
                    "inclusive_ir": edge["inclusive_ir"],
                    "position": edge.get("position"),
                }
                queue.append((callee_id, [*ids, callee_id], [*edges, edge_record]))
        return None

    caller = owner_edge["caller"]
    runner_path = positive_path(job["runner"], owner_edge["caller_id"])
    setup_paths = [
        positive_path(setup, owner_edge["caller_id"])
        for setup in job["setup_callers"]
    ]
    setup_paths = [path for path in setup_paths if path is not None]
    direct_runner = name_matches(caller, job["runner"])
    direct_setup = any(name_matches(caller, setup) for setup in job["setup_callers"])
    require(not (direct_runner and direct_setup),
            f"{parsed['path']}: owner caller matches timed and setup")
    require(not (runner_path is not None and setup_paths),
            f"{parsed['path']}: owner caller has timed and setup ancestry")
    if direct_runner or runner_path is not None:
        role = "timed"
        ancestry = runner_path or {
            "ancestor": job["runner"],
            "function_ids": [owner_edge["caller_id"]],
            "functions": [caller],
            "edges": [],
            "depth": 0,
            "max_depth": MAX_ANCESTRY_DEPTH,
            "direct_owner_caller": True,
        }
        setup_ancestry = None
    elif direct_setup or setup_paths:
        role = "setup"
        ancestry = setup_paths[0] if setup_paths else {
            "ancestor": next(setup for setup in job["setup_callers"]
                              if name_matches(caller, setup)),
            "function_ids": [owner_edge["caller_id"]],
            "functions": [caller],
            "edges": [],
            "depth": 0,
            "max_depth": MAX_ANCESTRY_DEPTH,
            "direct_owner_caller": True,
        }
        setup_ancestry = ancestry
    elif job["kind"] == "xls" and name_matches(caller, XLS_OWNER_WRAPPER):
        # This is an explicit diagnostic error rather than accepting a
        # wrapper without a positive path to the requested XLS runner.
        raise EvidenceError(
            f"{parsed['path']}: XLS wrapper has no positive runner ancestry"
        )
    else:
        raise EvidenceError(
            f"{parsed['path']}: owner caller {caller!r} has no allowed positive ancestry"
        )
    return {
        "role": role,
        "part": parsed["part"],
        "owner_id": owner_id,
        "owner_name": owner["name"],
        "owner_caller": caller,
        "owner_caller_id": owner_edge["caller_id"],
        "owner_calls": owner_edge["calls"],
        "owner_incoming": owner_edge,
        "owner_self_ir": owner["self_ir"],
        "owner_direct_ir": direct_ir,
        "summary_ir": parsed["summary"],
        "runner_ancestry": runner_path,
        "setup_ancestry": setup_ancestry,
        "role_validation": {
            "positive_owner_edge": True,
            "positive_runner_or_setup_ancestry": True,
            "role_not_inferred_from_dump_ordinal": True,
            "part_retained_as_historical_shape_only": True,
            "owner_edge_matches_summary": True,
            "owner_self_plus_direct_matches_summary": True,
        },
    }


def numbered_paths(folder: Path, name: str) -> list[tuple[int, Path]]:
    prefix = f"{name}.callgrind."
    found: list[tuple[int, Path]] = []
    for path in folder.glob(prefix + "*"):
        suffix = path.name[len(prefix):]
        if suffix.isdigit() and path.is_file() and not path.is_symlink():
            found.append((int(suffix), path))
    found.sort()
    require(found, f"{relative(folder / name)}: no numbered dumps")
    require([number for number, _ in found] == list(range(1, len(found) + 1)),
            f"{name}: numbered dumps are not contiguous")
    return found


def parse_assembly(stage: str) -> dict[str, Any]:
    folder = HERE / stage
    index_path = folder / "assembly-index.json"
    index = read_json(index_path)
    require(isinstance(index, dict), f"{relative(index_path)} is not an object")
    require(index.get("schema") == "ole2_0554_assembly_v1",
            f"{relative(index_path)} schema differs")
    require(index.get("plan_sha256") == sha256(PLAN_PATH),
            f"{relative(index_path)} plan binding differs")
    source_manifest = folder / "source-manifest.json"
    require(index.get("source_manifest_sha256") == sha256(source_manifest),
            f"{relative(index_path)} source binding differs")
    require(index.get("script_sha256") == sha256(INSPECT_PATH),
            f"{relative(index_path)} inspection script binding differs")
    binary_meta_path = folder / "binary-normal.json"
    binary_meta = read_json(binary_meta_path)
    require(index.get("binary_sha256") == binary_meta.get("sha256"),
            f"{relative(index_path)} binary hash differs from binary metadata")
    binary_path = binary_meta.get("path")
    require(isinstance(binary_path, str) and binary_path,
            f"{relative(binary_meta_path)} binary path is missing")
    rows = index.get("rows")
    require(isinstance(rows, list) and rows, f"{relative(index_path)} has no rows")
    parsed_rows: list[dict[str, Any]] = []
    for row in rows:
        require(isinstance(row, dict), f"{relative(index_path)} row is not an object")
        for key in ("name", "symbol", "address_hex", "size_bytes", "receipt_sha256"):
            require(key in row, f"{relative(index_path)} row lacks {key}")
        stdout_path = folder / f"{row['name']}.stdout"
        stderr_path = folder / f"{row['name']}.stderr"
        host_path = folder / f"{row['name']}.host.json"
        receipt_path = folder / f"{row['name']}.receipt.json"
        require(stdout_path.is_file() and not stdout_path.is_symlink(),
                f"missing assembly stdout {stdout_path}")
        require(stderr_path.is_file() and not stderr_path.is_symlink(),
                f"missing assembly stderr {stderr_path}")
        require(host_path.is_file() and not host_path.is_symlink(),
                f"missing assembly host observation {host_path}")
        receipt = read_json(receipt_path)
        require(receipt.get("execution_stage") == stage,
                f"{relative(receipt_path)} execution stage differs")
        require(receipt.get("binary_sha256") == index["binary_sha256"],
                f"{relative(receipt_path)} binary binding differs")
        require(receipt.get("plan_sha256") == index["plan_sha256"],
                f"{relative(receipt_path)} plan binding differs")
        require(receipt.get("source_manifest_sha256") == index["source_manifest_sha256"],
                f"{relative(receipt_path)} source binding differs")
        command = receipt.get("command")
        require(command == [
            "objdump", "-d", f"--disassemble={row['symbol']}", binary_path
        ], f"{relative(receipt_path)} symbol-bounded command differs")
        artifacts = receipt.get("artifacts")
        require(isinstance(artifacts, dict),
                f"{relative(receipt_path)} artifacts are missing")
        for artifact in (
            f"{row['name']}.host.json", f"{row['name']}.stdout", f"{row['name']}.stderr"
        ):
            artifact_path = folder / artifact
            require(artifacts.get(artifact) == sha256(artifact_path),
                    f"{relative(receipt_path)} artifact hash differs: {artifact}")
        require(sha256(receipt_path) == row["receipt_sha256"],
                f"assembly receipt hash differs for {row['name']}")
        parsed = RAW.parse_objdump(row, stdout_path)
        parsed["assembly_receipt"] = {
            "file": relative(receipt_path),
            "sha256": row["receipt_sha256"],
        }
        parsed_rows.append(parsed)

    static_by_address: dict[int, dict[str, Any]] = {}
    categories: dict[str, list[dict[str, Any]]] = {
        key: [] for key in ALL_MARKERS
    }
    for row in parsed_rows:
        row_categories = [
            key for key, markers in STATIC_MARKERS.items()
            if static_marker_matches(row["symbol"], key)
        ]
        row["target_categories"] = row_categories
        for key in row_categories:
            categories[key].append(row)
        for address, instruction in row["instructions"].items():
            previous = static_by_address.get(address)
            require(previous is None or previous["symbol"] == row["symbol"],
                    f"static instruction collision at {address:#x}")
            static_by_address[address] = {
                **instruction,
                "address": address,
                "symbol": row["symbol"],
                "function_start": row["address"],
                "function_size": row["size_bytes"],
                "target_categories": row_categories,
            }
    return {
        "index": index,
        "index_sha256": sha256(index_path),
        "binary_sha256": index["binary_sha256"],
        "source_manifest_sha256": index["source_manifest_sha256"],
        "rows": parsed_rows,
        "categories": categories,
        "static_by_address": static_by_address,
    }


def target_ids(parsed: dict[str, Any], target: str) -> list[int]:
    return sorted(
        function_id for function_id, function in parsed["functions"].items()
        if marker_matches(function.get("name", ""), TARGET_MARKERS[target])
    )


def target_edges(parsed: dict[str, Any], caller_target: str,
                 callee_target: str) -> list[dict[str, Any]]:
    callees = set(target_ids(parsed, callee_target))
    rows = []
    for parent_id, parent in parsed["functions"].items():
        if not marker_matches(parent.get("name", ""), TARGET_MARKERS[caller_target]):
            continue
        for edge in parent.get("edges", []):
            if (edge.get("callee_id") in callees
                    and edge.get("calls") is not None
                    and edge["calls"] > 0
                    and edge.get("inclusive_ir", 0) > 0):
                rows.append({
                    "caller_id": parent_id,
                    "caller": parent.get("name", ""),
                    "callee_id": edge["callee_id"],
                    "callee": parsed["functions"].get(edge["callee_id"], {}).get("name", ""),
                    "calls": edge["calls"],
                    "inclusive_ir": edge["inclusive_ir"],
                    "position": edge.get("position"),
                })
    rows.sort(key=lambda row: (row["caller"], row["callee"], row["callee_id"]))
    return rows


def dynamic_target(parsed: dict[str, Any], target: str, assembly: dict[str, Any],
                   bias: int) -> dict[str, Any]:
    """Map self instruction records for one target and one dump."""

    ids = target_ids(parsed, target)
    if not ids:
        return {
            "target": target,
            "markers": list(TARGET_MARKERS[target]),
            "state": "absent_or_inlined",
            "absence_reason": "no_named_function_record",
            "function_ids": [],
            "self_instruction_ir": 0,
            "mapped_instruction_ir": 0,
            "instruction_count": 0,
            "unique_static_instruction_count": 0,
            "functions": [],
            "mapping_complete": False,
            "interpretation": (
                "No named function record was present; owner/caller disassembly "
                "is required before inferring moved or eliminated work."
            ),
        }

    static_rows = assembly["categories"][target]
    static_addresses = {
        address for row in static_rows for address in row["instructions"]
    }
    functions: list[dict[str, Any]] = []
    all_records: list[dict[str, Any]] = []
    self_total = 0
    mapped_total = 0
    mapping_complete = bool(static_rows)
    absence_reason = None if static_rows else "no_matching_static_symbol"
    for function_id in ids:
        function = parsed["functions"][function_id]
        records: list[dict[str, Any]] = []
        function_self = int(function["self_ir"])
        function_mapped = 0
        for raw_address, ir in sorted(function["instruction_ir"].items()):
            static_address = raw_address - bias
            static = assembly["static_by_address"].get(static_address)
            if static is None:
                mapping_complete = False
                records.append({
                    "raw_address_hex": f"{raw_address:#x}",
                    "static_address_hex": f"{static_address:#x}",
                    "ir": int(ir),
                    "mapped": False,
                })
                continue
            if target not in static.get("target_categories", []):
                raise EvidenceError(
                    f"{parsed['path']}: {target} position {raw_address:#x} maps "
                    f"to {static['symbol']} outside its bounded symbol"
                )
            function_mapped += int(ir)
            record = {
                "raw_address_hex": f"{raw_address:#x}",
                "static_address_hex": f"{static_address:#x}",
                "offset_hex": f"{static_address - static['function_start']:#x}",
                "ir": int(ir),
                "mapped": True,
                "symbol": static["symbol"],
                "bytes_hex": static.get("bytes_hex", ""),
                "text": static.get("text"),
            }
            records.append(record)
            all_records.append(record)
        require(sum(function["instruction_ir"].values()) == function_self,
                f"{parsed['path']}: {target} function self/instruction Ir differs")
        self_total += function_self
        mapped_total += function_mapped
        functions.append({
            "function_id": function_id,
            "name": function["name"],
            "object": function.get("object"),
            "self_ir": function_self,
            "instruction_count": len(function["instruction_ir"]),
            "mapped_instruction_count": sum(1 for row in records if row["mapped"]),
            "mapped_instruction_ir": function_mapped,
            "mapping_complete": function_mapped == function_self,
            "instructions": records,
        })
    unique_static = len({row["static_address_hex"] for row in all_records if row["mapped"]})
    if not mapping_complete:
        absence_reason = absence_reason or "one_or_more_instruction_positions_unmapped"
    state = "mapped" if mapping_complete else "partial_static_mapping"
    return {
        "target": target,
        "markers": list(TARGET_MARKERS[target]),
        "state": state,
        "absence_reason": absence_reason,
        "function_ids": ids,
        "self_instruction_ir": self_total,
        "mapped_instruction_ir": mapped_total,
        "instruction_count": sum(len(function["instruction_ir"]) for function in
                                  (parsed["functions"][ident] for ident in ids)),
        "unique_static_instruction_count": unique_static,
        "functions": functions,
        "mapping_complete": mapping_complete,
        "interpretation": (
            "Exclusive self instruction Ir mapped to exact bounded objdump symbols."
            if mapping_complete else
            "Named work was present but static bounds were incomplete; inspect the "
            "owner/caller disassembly before interpreting code movement."
        ),
    }


def collect_raw_dumps(stage: str, job: dict[str, Any], plan: dict[str, Any],
                      assembly: dict[str, Any]) -> tuple[dict[str, Any], list[dict[str, Any]]]:
    receipt = validate_receipt(stage, job, plan)
    folder = HERE / stage
    paths = numbered_paths(folder, job["name"])
    expected = 5 if job["kind"] == "xls" else 6
    require(len(paths) == expected, f"{job['name']}: numbered dump count differs")
    parsed_dumps = [RAW.parse_raw_profile(path.resolve()) for _, path in paths]
    raw_addresses = {
        address
        for parsed in parsed_dumps
        for function in parsed["functions"].values()
        if any(marker_matches(function.get("name", ""), markers)
               for markers in TARGET_MARKERS.values())
        for address in function["instruction_ir"]
    }
    assembly_addresses = set(assembly["static_by_address"])
    require(raw_addresses, f"{job['name']}: no target instruction positions")
    # A single PIE relocation applies to all target functions.  Candidate
    # biases are checked against the full target address union, rather than a
    # hand-picked instruction or a caller-inclusive cost.
    first = min(raw_addresses)
    candidates = {
        first - static
        for static in assembly_addresses
        if all(raw - (first - static) in assembly_addresses for raw in raw_addresses)
    }
    require(len(candidates) == 1,
            f"{job['name']}: expected one relocation bias, got {sorted(candidates)}")
    bias = next(iter(candidates))
    dumps: list[dict[str, Any]] = []
    for (number, path), parsed in zip(paths, parsed_dumps):
        # The suffix only identifies the retained artifact's sequence.  The
        # parser's part record is preserved below and role classification uses
        # positive ancestry, so a fresh setup/timed shape cannot be silently
        # inferred from ``.callgrind.N``.
        role = parse_dump_role(parsed, job, number)
        targets = {
            target: dynamic_target(parsed, target, assembly, bias)
            for target in TARGET_MARKERS
        }
        decoder_edges = target_edges(parsed, "parse_directory_entry", "decode_utf16le")
        scalar_edges = target_edges(parsed, "parse_directory_entry", "format_clsid")
        dumps.append({
            "number": number,
            "part": parsed["part"],
            "file": relative(path),
            "sha256": sha256(path),
            "role": role["role"],
            "owner": {
                "id": role["owner_id"],
                "name": role["owner_name"],
                "caller": role["owner_caller"],
                "caller_id": role["owner_caller_id"],
                "calls": role["owner_calls"],
                "incoming": role["owner_incoming"],
                "self_ir": role["owner_self_ir"],
                "direct_ir": role["owner_direct_ir"],
                "summary_ir": role["summary_ir"],
            },
            "runner_ancestry": role["runner_ancestry"],
            "setup_ancestry": role["setup_ancestry"],
            "targets": targets,
            "direct_edges": {
                "parse_directory_entry_to_decode_utf16le": {
                    "edges": decoder_edges,
                    "calls": sum(edge["calls"] for edge in decoder_edges),
                    "inclusive_ir": sum(edge["inclusive_ir"] for edge in decoder_edges),
                },
                "parse_directory_entry_to_format_clsid": {
                    "edges": scalar_edges,
                    "calls": sum(edge["calls"] for edge in scalar_edges),
                    "inclusive_ir": sum(edge["inclusive_ir"] for edge in scalar_edges),
                },
            },
            "accounting": {
                "instruction_ir_scope": "exclusive function self instruction Ir",
                "inclusive_cfn_ir_added": False,
                "inclusive_parent_ir_added": False,
                "calls_used_as_instruction_cost": False,
            },
            "validation": {
                **role["role_validation"],
                "positions_are_absolute_instr": parsed["positions"] == ["instr"],
                "self_instruction_sums_match": all(
                    sum(function["instruction_ir"].values()) == function["self_ir"]
                    for function in parsed["functions"].values()
                ),
            },
        })
    dumps.sort(key=lambda dump: dump["number"])
    return {
        "name": job["name"],
        "repeat": job["repeat"],
        "group": job["group"],
        "kind": job["kind"],
        "shape": job["shape"],
        "owner": job["owner"],
        "receipt": receipt,
        "relocation_bias": bias,
    }, dumps


def aggregate_job(dumps: list[dict[str, Any]], role: str = "timed") -> dict[str, Any]:
    result: dict[str, Any] = {}
    selected = [dump for dump in dumps if dump["role"] == role]
    for target in TARGET_MARKERS:
        rows = [dump["targets"][target] for dump in selected]
        static_addresses = {
            instruction["static_address_hex"]
            for row in rows
            for function in row["functions"]
            for instruction in function["instructions"]
            if instruction.get("mapped")
        }
        result[target] = {
            "dump_count": len(rows),
            "states": [row["state"] for row in rows],
            "self_instruction_ir": sum(row["self_instruction_ir"] for row in rows),
            "mapped_instruction_ir": sum(row["mapped_instruction_ir"] for row in rows),
            "instruction_count": sum(row["instruction_count"] for row in rows),
            # This is a union over the selected timed dumps in this job.  A
            # repeated instruction address is counted once here, while each
            # dump retains its own per-dump unique count above.
            "unique_static_instruction_count": len(static_addresses),
            "mapping_complete_all_dumps": all(row["mapping_complete"] for row in rows),
        }
    return result


def analyze_stage(stage: str, plan: dict[str, Any]) -> dict[str, Any]:
    require(stage in STAGES, f"unsupported stage {stage}")
    assembly = parse_assembly(stage)
    profiles: list[dict[str, Any]] = []
    for job in profile_jobs(plan):
        profile, dumps = collect_raw_dumps(stage, job, plan, assembly)
        setup = [dump for dump in dumps if dump["role"] == "setup"]
        timed = [dump for dump in dumps if dump["role"] == "timed"]
        expected_setup = 0 if job["kind"] == "xls" else 1
        require(len(setup) == expected_setup,
                f"{job['name']}: actual positive-ancestry setup count differs")
        require(len(timed) == int(plan["profile"]["samples"]),
                f"{job['name']}: actual positive-ancestry timed count differs")
        profiles.append({
            **profile,
            "dumps": dumps,
            "setup_dump_count": len(setup),
            "timed_dump_count": len(timed),
            "timed_aggregate": aggregate_job(dumps),
            "validation": {
                "roles_from_positive_ancestry": True,
                "role_not_from_dump_ordinal": all(
                    dump["validation"]["role_not_inferred_from_dump_ordinal"]
                    for dump in dumps
                ),
                "setup_excluded_from_timed_aggregates": True,
                "all_target_mappings_explicit": True,
                "inclusive_ir_not_added_to_instruction_totals": all(
                    not dump["accounting"]["inclusive_cfn_ir_added"]
                    and not dump["accounting"]["inclusive_parent_ir_added"]
                    for dump in dumps
                ),
            },
        })
    candidate_decoder_exact = (
        all(
            dump["direct_edges"]["parse_directory_entry_to_decode_utf16le"]["calls"] == 1
            for profile in profiles
            for dump in profile["dumps"]
            if dump["role"] == "timed"
        )
        if stage == "candidate" else None
    )
    return {
        "schema": "ole2_name_handoff_0554_instruction_analysis_v1",
        "status": "pass",
        "stage": stage,
        "scope": plan["scope"],
        "performance_claim": "diagnostic-only",
        "plan": relative(PLAN_PATH),
        "plan_sha256": sha256(PLAN_PATH),
        "assembly": {
            "index": relative(HERE / stage / "assembly-index.json"),
            "index_sha256": assembly["index_sha256"],
            "binary_sha256": assembly["binary_sha256"],
            "source_manifest_sha256": assembly["source_manifest_sha256"],
            "row_count": len(assembly["rows"]),
            "static_target_row_counts": {
                target: len(assembly["categories"][target]) for target in ALL_MARKERS
            },
            "relocation_biases": sorted({profile["relocation_bias"] for profile in profiles}),
            "target_symbols": {
                target: [row["symbol"] for row in assembly["categories"][target]]
                for target in ALL_MARKERS
            },
        },
        "profiles": profiles,
        "mechanism": {
            "target_instruction_totals_are_exclusive_self_ir": True,
            "dynamic_decoder_edge_is_incoming_parse_to_decode": True,
            "scalar_clsid_edge_is_retained_as_control": True,
            "candidate_decoder_calls_exactly_one_checked_in_profile_report": candidate_decoder_exact,
            "static_mapping_required_for_work_removal_interpretation": True,
            "symbol_absence_is_not_zero_work": True,
        },
        "helpers": {
            "change-0547/instruction_analysis.py": sha256(RAW_PARSER_PATH),
            "run.py": sha256(RUN_PATH),
            "inspect_assembly.py": sha256(INSPECT_PATH),
            "instruction_analysis.py": sha256(SCRIPT_PATH),
        },
        "validation": {
            "assembly_receipts_and_stdout_hashed": True,
            "plan_binary_and_source_bindings_checked": True,
            "all_profile_receipts_and_numbered_dumps_hashed": True,
            "absolute_positions_relocated_to_static_addresses": True,
            "bounded_symbol_addresses_and_sizes_checked": True,
            "positive_owner_edges_and_ancestry_checked": True,
            "setup_parts_retained_but_excluded": True,
            "no_inclusive_parent_sum": True,
            "no_native_or_allocation_claim": True,
        },
        "limitations": [
            "The report maps only self instruction Ir within symbols retained by the stage assembly index; inlined work requires the selected owner/caller disassembly.",
            "Callgrind Ir and dynamic edge counts establish mechanism evidence, not native latency, RSS, allocation, hardware, or scaling behavior.",
            "A zero or absent named function record is reported as absent_or_inlined and is never interpreted as eliminated work.",
        ],
    }


def metric(baseline: int, candidate: int) -> dict[str, int | float | None]:
    delta = candidate - baseline
    return {
        "baseline": baseline,
        "candidate": candidate,
        "delta": delta,
        "delta_percent": None if baseline == 0 else 100.0 * delta / baseline,
    }


def profile_key(profile: dict[str, Any]) -> tuple[int, str]:
    return int(profile["repeat"]), str(profile["group"])


def compare_stages(baseline: dict[str, Any], candidate: dict[str, Any]) -> dict[str, Any]:
    b_profiles = {profile_key(profile): profile for profile in baseline["profiles"]}
    c_profiles = {profile_key(profile): profile for profile in candidate["profiles"]}
    require(set(b_profiles) == set(c_profiles), "stage profile matrices differ")
    rows: list[dict[str, Any]] = []
    for key in sorted(b_profiles):
        bp, cp = b_profiles[key], c_profiles[key]
        targets: dict[str, Any] = {}
        for target in TARGET_MARKERS:
            b = bp["timed_aggregate"][target]
            c = cp["timed_aggregate"][target]
            targets[target] = {
                field: metric(int(b[field]), int(c[field]))
                for field in (
                    "self_instruction_ir", "mapped_instruction_ir",
                    "instruction_count", "unique_static_instruction_count",
                )
            }
            targets[target]["baseline_states"] = b["states"]
            targets[target]["candidate_states"] = c["states"]
            targets[target]["baseline_mapping_complete"] = b["mapping_complete_all_dumps"]
            targets[target]["candidate_mapping_complete"] = c["mapping_complete_all_dumps"]
        b_decoder = [
            dump["direct_edges"]["parse_directory_entry_to_decode_utf16le"]["calls"]
            for dump in bp["dumps"] if dump["role"] == "timed"
        ]
        c_decoder = [
            dump["direct_edges"]["parse_directory_entry_to_decode_utf16le"]["calls"]
            for dump in cp["dumps"] if dump["role"] == "timed"
        ]
        b_scalar = [
            dump["direct_edges"]["parse_directory_entry_to_format_clsid"]["calls"]
            for dump in bp["dumps"] if dump["role"] == "timed"
        ]
        c_scalar = [
            dump["direct_edges"]["parse_directory_entry_to_format_clsid"]["calls"]
            for dump in cp["dumps"] if dump["role"] == "timed"
        ]
        rows.append({
            "repeat": key[0],
            "group": key[1],
            "targets": targets,
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
    many_small = [row for row in rows if row["group"] == "cfb-many-small"]
    return {
        "rows": rows,
        "mechanism_gate": {
            "candidate_decoder_edge_exactly_one_each_timed_dump": all(
                row["decoder_calls_per_timed_dump"]["candidate_all_exactly_one"]
                for row in rows
            ),
            "many_small_decoder_self_instruction_ir_decreases_each_repeat": all(
                row["targets"]["decode_utf16le"]["self_instruction_ir"]["delta"] < 0
                for row in many_small
            ),
            "many_small_decoder_mapping_complete_both_stages": all(
                row["targets"]["decode_utf16le"]["baseline_mapping_complete"]
                and row["targets"]["decode_utf16le"]["candidate_mapping_complete"]
                for row in many_small
            ),
            "format_clsid_call_vector_preserved": all(
                row["format_clsid_calls_per_timed_dump"]["same_call_vector"]
                for row in rows
            ),
            "format_clsid_mapping_complete_both_stages": all(
                row["targets"]["format_clsid"]["baseline_mapping_complete"]
                and row["targets"]["format_clsid"]["candidate_mapping_complete"]
                for row in rows
            ),
            "scalar_and_decoder_instruction_rows_are_separate": True,
        },
        "interpretation": (
            "Self instruction Ir is compared only after each raw position maps to "
            "the exact bounded symbol.  Inclusive parent/callee Ir is retained "
            "separately and is never added to these totals."
        ),
    }


def compare(plan: dict[str, Any]) -> dict[str, Any]:
    baseline = analyze_stage("baseline", plan)
    candidate = analyze_stage("candidate", plan)
    comparison = compare_stages(baseline, candidate)
    return {
        "schema": "ole2_name_handoff_0554_instruction_comparison_v1",
        "status": "pass",
        "scope": plan["scope"],
        "performance_claim": "diagnostic-only",
        "plan": relative(PLAN_PATH),
        "plan_sha256": sha256(PLAN_PATH),
        "stages": {"baseline": baseline, "candidate": candidate},
        "comparison": comparison,
        "validation": {
            "matched_stage_profiles_compared": True,
            "per_dump_decoder_and_scalar_edges_compared": True,
            "self_instruction_ir_only": True,
            "static_mapping_precedes_work_interpretation": True,
            "no_native_or_allocation_claim": True,
        },
    }


def write_report(document: dict[str, Any], output: Path) -> None:
    data = (json.dumps(document, indent=2, sort_keys=True) + "\n").encode("utf-8")
    if output.exists():
        require(output.is_file() and not output.is_symlink(),
                f"refusing non-regular output {output}")
        require(output.read_bytes() == data,
                f"refusing to overwrite non-identical output {output}")
        return
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_bytes(data)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("output", nargs="?", type=Path)
    parser.add_argument("--output", dest="output_option", type=Path)
    parser.add_argument("--stage", choices=STAGES, default="baseline")
    parser.add_argument("--compare", action="store_true")
    args = parser.parse_args(argv)
    if args.output is not None and args.output_option is not None:
        parser.error("provide output positionally or with --output, not both")
    output = (args.output_option or args.output or
              (HERE / "instruction-comparison.json" if args.compare else
               HERE / args.stage / "instruction-analysis.json"))
    try:
        plan = load_plan()
        document = compare(plan) if args.compare else analyze_stage(args.stage, plan)
        write_report(document, output)
    except (EvidenceError, OSError, TypeError, ValueError, KeyError) as error:
        print(f"instruction_analysis.py: error: {error}", file=sys.stderr)
        return 2
    print(output)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
