#!/usr/bin/env python3
"""Validate and attribute the matched 0555 OLE2 Callgrind profiles.

This is a read-only evidence consumer.  It reuses the retained 0554 raw
Callgrind parser, but owns the 0555 plan, receipt bindings, role
classification, and comparison schema.  Every CFB setup dump is classified
from a positive setup ancestry edge; dump numbering is retained as an
artifact identity and is never used to decide whether a dump was setup or
timed.

The physical-marker experiment moves bookkeeping between ``load_fat``, the
directory-chain validation callers, and the final physical reconciliation
leaf.  The report therefore keeps selected-owner inclusive/self/direct Ir,
each requested target's inclusive/self/direct Ir, and explicit inline/absent
states in separate fields.  A missing or inlined symbol has no synthetic zero
cost: its work remains ``indeterminate_assembly_required`` until the matching
static instruction mapping is reviewed.  No native, allocation, or adoption
claim is made here.
"""

from __future__ import annotations

import argparse
import datetime
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
LOCK_BINDING = HERE / "workspace-lock.json"
LOCK_COPY = HERE / "workspace-Cargo.lock"
STAGES = ("baseline", "candidate")
TARGET = Path("/home/zhuhe/litchi-goal-0555-target")
OLD_ANALYZER_PATH = HERE.parent / "change-0554" / "analyze_profiles.py"

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
XLS_OWNER_WRAPPER = (
    "litchi_xls::workbook::source::SourceBackedWorkbook::from_read_at"
)
CFB_SHAPES = ("tiny", "many-small", "few-large")
XLS_CASE = "xls_owned_source_open_one_cell"
CFB_CASE = "cfb_open"
MAX_ANCESTRY_DEPTH = 8

# These are diagnostic attribution targets, not additive cost buckets.  The
# target names are private implementation names in the measured CFB binary;
# exact positive incoming edges are retained wherever the compiler emits one.
ATTRIBUTION_TARGETS = {
    "load_fat": "litchi_cfb::file::OleFile<R>::load_fat",
    "claim_sector": "litchi_cfb::file::OleFile<R>::claim_sector",
    "validate_stream_allocations": (
        "litchi_cfb::file::OleFile<R>::validate_stream_allocations"
    ),
    "collect_exact": "litchi_cfb::file::SectorChainScratch::collect_exact",
    "physical_reconciliation": (
        "litchi_cfb::file::OleFile<R>::validate_physical_sector_layout"
    ),
}
INLINE_CALLER_MARKERS = (
    "::claim_chain",
    "::validate_stream_allocations",
    "::load_fat",
    "::open",
)

PROFILE_FLAGS = (
    "--dump-instr=yes",
    "--dump-line=no",
    "--compress-pos=no",
    "--collect-jumps=yes",
)
REQUIRED_PROFILE_ARTIFACTS = ("host.json", "json", "catalog.json", "stdout",
                              "stderr", "callgrind")
SHA256_RE = re.compile(r"^[0-9a-f]{64}$")
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
        return path.resolve().relative_to(HERE.resolve()).as_posix()
    except ValueError as error:
        raise EvidenceError(f"path is outside evidence directory: {path}") from error


def check_hash(value: Any, label: str) -> str:
    require(isinstance(value, str) and SHA256_RE.fullmatch(value) is not None,
            f"{label}: expected lowercase SHA-256")
    return value


def timestamp(value: Any, label: str) -> datetime.datetime:
    require(isinstance(value, str), f"{label}: timestamp is missing")
    try:
        parsed = datetime.datetime.fromisoformat(value)
    except ValueError as error:
        raise EvidenceError(f"{label}: timestamp is malformed") from error
    require(parsed.tzinfo is not None, f"{label}: timestamp has no timezone")
    return parsed


def load_prior_analyzer() -> Any:
    """Load only the immutable 0554 parser module; never invoke its CLI."""

    require(OLD_ANALYZER_PATH.is_file(),
            f"missing retained 0554 profile analyzer: {OLD_ANALYZER_PATH}")
    spec = importlib.util.spec_from_file_location(
        "litchi_0555_retained_0554_profiles", OLD_ANALYZER_PATH
    )
    require(spec is not None and spec.loader is not None,
            f"cannot load retained analyzer: {OLD_ANALYZER_PATH}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    # The parser uses HERE only for diagnostics.  Rebinding its module-local
    # path keeps emitted error paths relative to this campaign and does not
    # mutate the retained source file or any old artifact.
    module.HERE = HERE
    module.RAW.HERE = HERE
    return module


PRIOR = load_prior_analyzer()
RAW = PRIOR.RAW


def plan_data() -> dict[str, Any]:
    plan = read_json(PLAN_PATH)
    require(isinstance(plan, dict), "plan is not an object")
    require(plan.get("schema") == "ole2_physical_marker_0555_plan_v1",
            "plan schema differs")
    require(plan.get("scope") == (
        "Matched OLE2 physical-sector marker accounting experiment across the "
        "FAT, DIFAT, Directory, MiniFAT, MiniStream, and RegularStream roles; "
        "no public API or semantic format change"
    ), "plan scope differs")
    require(plan.get("candidate_files") == ["crates/litchi-cfb/src/file.rs"],
            "candidate file scope differs")
    require(plan.get("cpu") == 2, "profile CPU differs")
    require(plan.get("owned_paths") == [str(TARGET)], "owned target differs")
    require(plan.get("status") ==
            "frozen before build, capture, candidate application, and result observation",
            "plan is not frozen before capture")
    groups = plan.get("groups")
    require(isinstance(groups, dict), "plan groups are missing")
    require(XLS_CASE in groups.get("xls", {}).get("cases", []),
            "profile XLS case is absent")
    require(groups.get("cfb", {}).get("cases") == [CFB_CASE],
            "CFB profile case differs")
    require(tuple(groups.get("cfb", {}).get("shapes", ())) == CFB_SHAPES,
            "CFB profile shapes differ")
    require(groups.get("cfb", {}).get("payload") == "incompressible",
            "CFB profile payload differs")

    profile = plan.get("profile")
    require(isinstance(profile, dict), "profile plan is missing")
    require(profile.get("repeats") == 2 and profile.get("warmup") == 0
            and profile.get("samples") == 5,
            "profile sample matrix differs")
    require(profile.get("jobs") == [
        "xls-owned", "cfb-tiny", "cfb-many-small", "cfb-few-large"
    ], "profile job matrix differs")
    require(profile.get("xls_owner") == XLS_OWNER, "XLS owner differs")
    require(profile.get("cfb_owner") == CFB_OWNER, "CFB owner differs")
    require(profile.get("instruction_flags") == list(PROFILE_FLAGS),
            "Callgrind instruction flags differ")

    native = plan.get("native")
    allocation = plan.get("allocation")
    require(isinstance(native, dict) and native.get("repeats") == 2
            and native.get("samples") == 1000 and native.get("warmup") == 20,
            "native plan differs")
    require(isinstance(allocation, dict) and allocation.get("repeats") == 2
            and allocation.get("samples") == 30 and allocation.get("warmup") == 3,
            "allocation plan differs")
    return plan


def profile_jobs(plan: dict[str, Any]) -> list[dict[str, Any]]:
    profile = plan["profile"]
    jobs: list[dict[str, Any]] = []
    for repeat in range(1, profile["repeats"] + 1):
        jobs.append({
            "name": f"profile-r{repeat}-xls-owned",
            "repeat": repeat,
            "group": "xls-owned",
            "kind": "xls",
            "shape": None,
            "owner": profile["xls_owner"],
            "runner": XLS_RUNNER,
            "setup_callers": (),
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
                "setup_callers": CFB_SETUP_CALLERS,
                "selection": {
                    "cases": [CFB_CASE],
                    "shapes": [shape],
                    "payload": "incompressible",
                },
                "samples": profile["samples"],
                "warmup": profile["warmup"],
            })
    return jobs


def expected_execution_stage(stage: str, repeat: int) -> str:
    # The second baseline output leg retains the baseline binary while it is
    # run after candidate source installation.  Its receipt must say so.
    return "candidate" if stage == "baseline" and repeat == 2 else stage


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


def validate_host(path: Path, label: str) -> None:
    value = read_json(path)
    require(isinstance(value, dict), f"{label}: host observation is not an object")
    timestamp(value.get("observed_utc"), f"{label}.observed_utc")
    processes = value.get("compiler_processes")
    require(isinstance(processes, list), f"{label}: compiler observations are missing")
    for index, process in enumerate(processes):
        require(isinstance(process, dict), f"{label}: process {index} is malformed")
        require(isinstance(process.get("pid"), int)
                and not isinstance(process.get("pid"), bool)
                and process["pid"] > 0,
                f"{label}: process {index} pid is invalid")
        require(process.get("comm") in {"cargo", "rustc"},
                f"{label}: process {index} command is unexpected")
        require(isinstance(process.get("cwd"), str) and process["cwd"],
                f"{label}: process {index} cwd is missing")
    require(value.get("scope") ==
            "Accessible compiler processes; no host quiescence guarantee",
            f"{label}: host scope differs")


def validate_artifacts(folder: Path, receipt: dict[str, Any],
                       expected: set[str], label: str) -> None:
    artifacts = receipt.get("artifacts")
    require(isinstance(artifacts, dict), f"{label}: artifact map is missing")
    require(set(artifacts) == expected,
            f"{label}: artifact set differs; missing={sorted(expected - set(artifacts))}, "
            f"extra={sorted(set(artifacts) - expected)}")
    for filename, digest in artifacts.items():
        require(isinstance(filename, str) and Path(filename).name == filename,
                f"{label}: artifact is not stage-local: {filename!r}")
        check_hash(digest, f"{label}/{filename}")
        path = folder / filename
        require(path.is_file() and not path.is_symlink(),
                f"{label}: artifact is missing: {filename}")
        require(sha256(path) == digest, f"{label}: artifact hash differs: {filename}")
        if filename.endswith(".host.json"):
            validate_host(path, f"{label}/{filename}")


def validate_lock(receipt: dict[str, Any], label: str) -> None:
    if LOCK_BINDING.is_file() and LOCK_COPY.is_file():
        binding = read_json(LOCK_BINDING)
        require(isinstance(binding, dict)
                and binding.get("path") == "Cargo.lock"
                and binding.get("scope") ==
                "Ignored workspace lock additionally bound on every child; perf harness lock is in source manifest",
                f"{label}: workspace lock binding differs")
        lock_hash = check_hash(binding.get("sha256"), f"{label} lock hash")
        require(lock_hash == sha256(LOCK_COPY), f"{label}: lock copy hash differs")
        require(receipt.get("workspace_lock_sha256") == lock_hash,
                f"{label}: receipt lock hash differs")
        require(receipt.get("workspace_lock_binding_sha256") == sha256(LOCK_BINDING),
                f"{label}: receipt lock binding differs")


def validate_receipt(stage: str, job: dict[str, Any], plan: dict[str, Any]) -> dict[str, Any]:
    folder = HERE / stage
    path = folder / f"{job['name']}.receipt.json"
    receipt = read_json(path)
    require(isinstance(receipt, dict), f"{relative(path)} is not an object")
    label = relative(path)
    require(receipt.get("schema") == "ole2_0555_run_receipt_v1",
            f"{label}: receipt schema differs")
    require(receipt.get("exit_code") == 0, f"{label}: child failed")
    timestamp(receipt.get("start_utc"), f"{label}.start_utc")
    timestamp(receipt.get("end_utc"), f"{label}.end_utc")
    require(receipt["start_utc"] < receipt["end_utc"],
            f"{label}: receipt interval is not positive")
    execution = expected_execution_stage(stage, job["repeat"])
    require(receipt.get("stage") == stage and receipt.get("execution_stage") == execution,
            f"{label}: output/execution stage binding differs")
    manifest_path = folder / "source-manifest.json"
    execution_manifest = HERE / execution / "source-manifest.json"
    require(receipt.get("source_manifest_sha256") == sha256(manifest_path),
            f"{label}: output source manifest differs")
    require(receipt.get("execution_manifest_sha256") == sha256(execution_manifest),
            f"{label}: execution source manifest differs")
    require(receipt.get("plan_sha256") == sha256(PLAN_PATH),
            f"{label}: plan binding differs")
    require(receipt.get("script_sha256") == sha256(RUN_PATH),
            f"{label}: driver binding differs")
    validate_lock(receipt, label)

    command = receipt.get("command")
    require(isinstance(command, list) and all(isinstance(item, str) for item in command),
            f"{label}: command is not argv")
    require(command[:7] == [
        "taskset", "-c", str(plan["cpu"]), "valgrind", "--vgdb=no",
        "--tool=callgrind", "--collect-atstart=no",
    ], f"{label}: Callgrind command prefix differs")
    require(command.count("--vgdb=no") == 1, f"{label}: vgdb flag repeats")
    for flag in PROFILE_FLAGS:
        require(flag in command and command.count(flag) == 1,
                f"{label}: profile flag differs: {flag}")
    require(f"--toggle-collect={job['owner']}" in command
            and f"--zero-before={job['owner']}" in command
            and f"--dump-after={job['owner']}" in command,
            f"{label}: owner controls differ")
    require(option(command, "--callgrind-out-file") ==
            str(folder / f"{job['name']}.callgrind"),
            f"{label}: Callgrind output path differs")
    require(option(command, "--case") == ",".join(job["selection"]["cases"]),
            f"{label}: case differs")
    require(option(command, "--warmup") == str(job["warmup"]),
            f"{label}: warmup differs")
    require(option(command, "--samples") == str(job["samples"]),
            f"{label}: samples differs")
    require(option(command, "--json") == str(folder / f"{job['name']}.json"),
            f"{label}: JSON output differs")
    require(option(command, "--corpus-manifest") ==
            str(folder / f"{job['name']}.catalog.json"),
            f"{label}: catalog output differs")
    if job["shape"] is None:
        require("--shape" not in command and "--payload" not in command,
                f"{label}: XLS command carries CFB selectors")
    else:
        require(option(command, "--shape") == job["shape"],
                f"{label}: CFB shape differs")
        require(option(command, "--payload") == "incompressible",
                f"{label}: CFB payload differs")

    count = 5 if job["kind"] == "xls" else 6
    expected = {
        f"{job['name']}.{suffix}" for suffix in REQUIRED_PROFILE_ARTIFACTS
    }
    expected.update(f"{job['name']}.callgrind.{number}"
                    for number in range(1, count + 1))
    validate_artifacts(folder, receipt, expected, label)
    binary_hash = check_hash(receipt.get("binary_sha256"), f"{label} binary hash")
    return {
        "path": relative(path),
        "sha256": sha256(path),
        "execution_stage": execution,
        "execution_manifest_sha256": receipt["execution_manifest_sha256"],
        "source_manifest_sha256": receipt["source_manifest_sha256"],
        "binary_sha256": binary_hash,
        "command": command,
        "artifacts": {key: artifacts for key, artifacts in
                       sorted(receipt["artifacts"].items())},
    }


def validate_profile_report(folder: Path, job: dict[str, Any],
                            binary_hash: str) -> dict[str, Any]:
    path = folder / f"{job['name']}.json"
    value = read_json(path)
    require(isinstance(value, dict), f"{relative(path)}: report is not an object")
    rows = value.get("results")
    require(isinstance(rows, list) and len(rows) == 1,
            f"{relative(path)}: expected one result")
    row = rows[0]
    require(isinstance(row, dict) and row.get("case") == job["selection"]["cases"][0],
            f"{relative(path)}: case differs")
    if job["shape"] is not None:
        require(row.get("corpus", {}).get("shape") == job["shape"],
                f"{relative(path)}: corpus shape differs")
    identity = value.get("binary_identity")
    require(isinstance(identity, dict) and identity.get("binary_sha256") == binary_hash,
            f"{relative(path)}: binary identity differs")
    return {
        "path": relative(path),
        "sha256": sha256(path),
        "identity": {
            "case": row.get("case"),
            "corpus": row.get("corpus"),
            "sink": row.get("sink"),
            "source": row.get("source"),
            "output_sha256": row.get("output_sha256"),
        },
    }


def matches(actual: str, expected: str) -> bool:
    return actual == expected or actual.endswith("::" + expected.rsplit("::", 1)[-1]) \
        or expected in actual


def positive_path(parsed: dict[str, Any], target_id: int,
                  ancestor: str) -> dict[str, Any] | None:
    functions = parsed["functions"]
    starts = sorted(function_id for function_id, function in functions.items()
                    if matches(function.get("name", ""), ancestor))
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
                "functions": [functions.get(item, {}).get("name", "") for item in ids],
                "edges": edges,
                "depth": len(edges),
                "max_depth": MAX_ANCESTRY_DEPTH,
            }
        if len(edges) >= MAX_ANCESTRY_DEPTH:
            continue
        current = functions.get(current_id)
        if current is None:
            continue
        candidates = [edge for edge in current.get("edges", [])
                      if edge.get("calls") is not None
                      and edge["calls"] >= 0
                      and edge.get("inclusive_ir", 0) > 0]
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


def classify_owner(parsed: dict[str, Any], job: dict[str, Any],
                   label: str) -> dict[str, Any]:
    """Classify the selected owner from positive ancestry, never ordinal."""

    edge = parsed["owner_incoming"]
    caller = edge["caller"]
    caller_id = edge["caller_id"]
    direct_runner = matches(caller, job["runner"])
    setup_matches = [expected for expected in job["setup_callers"]
                     if matches(caller, expected)]
    require(not (direct_runner and setup_matches),
            f"{label}: owner caller matches timed and setup roles")
    runner_path = None
    setup_path = None
    if direct_runner:
        role = "timed"
        runner_path = {
            "ancestor": job["runner"],
            "function_ids": [caller_id],
            "functions": [caller],
            "edges": [],
            "depth": 0,
            "max_depth": MAX_ANCESTRY_DEPTH,
            "direct_owner_caller": True,
        }
    elif setup_matches:
        role = "setup"
        setup_path = {
            "ancestor": setup_matches[0],
            "function_ids": [caller_id],
            "functions": [caller],
            "edges": [],
            "depth": 0,
            "max_depth": MAX_ANCESTRY_DEPTH,
            "direct_owner_caller": True,
        }
    else:
        runner_path = positive_path(parsed, caller_id, job["runner"])
        setup_paths = [positive_path(parsed, caller_id, expected)
                       for expected in job["setup_callers"]]
        setup_path = next((path for path in setup_paths if path is not None), None)
        require(not (runner_path is not None and setup_path is not None),
                f"{label}: owner caller is reachable from timed and setup paths")
        if runner_path is not None:
            role = "timed"
        elif setup_path is not None:
            role = "setup"
        elif job["kind"] == "xls" and matches(caller, XLS_OWNER_WRAPPER):
            raise EvidenceError(f"{label}: XLS wrapper has no positive runner ancestry")
        else:
            raise EvidenceError(f"{label}: owner caller has no allowed positive ancestry")
    require(edge["calls"] == 1, f"{label}: selected owner calls are not one")
    return {
        "role": role,
        "owner_caller": caller,
        "owner_caller_id": caller_id,
        "owner_calls": edge["calls"],
        "owner_incoming": edge,
        "runner_ancestry": runner_path,
        "setup_ancestry": setup_path,
        "validation": {
            "positive_owner_edge": True,
            "positive_ancestry_classified": True,
            "single_owner_call": True,
            "role_not_inferred_from_dump_ordinal": True,
        },
    }


def target_ids(parsed: dict[str, Any], target: str) -> list[int]:
    return sorted(function_id for function_id, function in parsed["functions"].items()
                  if target in function.get("name", ""))


def caller_context(parsed: dict[str, Any], label: str) -> list[dict[str, Any]]:
    """Retain possible inline callers without assigning their cost to claim_sector."""

    functions = parsed["functions"]
    rows = []
    for function_id, function in functions.items():
        name = function.get("name", "")
        if not any(marker in name for marker in INLINE_CALLER_MARKERS):
            continue
        incoming = [
            {
                "caller_id": parent_id,
                "caller": parent.get("name", ""),
                "callee_id": function_id,
                "callee": name,
                "calls": edge.get("calls"),
                "inclusive_ir": edge.get("inclusive_ir"),
                "line": edge.get("line"),
            }
            for parent_id, parent in functions.items()
            for edge in parent.get("edges", [])
            if edge.get("callee_id") == function_id
            and edge.get("calls") is not None
            and edge["calls"] > 0
            and edge.get("inclusive_ir", 0) > 0
        ]
        rows.append({
            "function_id": function_id,
            "function": name,
            "self_ir": function.get("self_ir"),
            "direct_ir": sum(edge.get("inclusive_ir", 0)
                              for edge in function.get("edges", [])),
            "inclusive_ir": (sum(edge["inclusive_ir"] for edge in incoming)
                             if incoming else None),
            "calls": (sum(edge["calls"] for edge in incoming)
                      if incoming else None),
            "incoming_edges": sorted(incoming, key=lambda item: (
                item["caller"], item["caller_id"], item["line"] or 0,
            )),
        })
    rows.sort(key=lambda item: (item["function"], item["function_id"]))
    return rows


def attribution(parsed: dict[str, Any], key: str, target: str,
                 label: str) -> dict[str, Any]:
    functions = parsed["functions"]
    ids = target_ids(parsed, target)
    if not ids:
        return {
            "target": target,
            "function_ids": [],
            "state": "inlined_or_absent",
            "absence_reason": "function_absent",
            "self_ir": None,
            "direct_ir": None,
            "inclusive_ir": None,
            "calls": None,
            "incoming_edge_count": None,
            "incoming_edges": [],
            "inline_fallback": {
                "status": "indeterminate_assembly_required",
                "candidate_callers": caller_context(parsed, label)
                if key == "claim_sector" else [],
                "value_substituted_for_target": False,
            },
            "validation": {
                "target_function_present": False,
                "positive_incoming_edge": False,
                "missing_is_not_zero_work": True,
                "assembly_required": True,
            },
        }

    selected = set(ids)
    incoming = []
    for parent_id, parent in functions.items():
        for edge in parent.get("edges", []):
            if edge.get("callee_id") in selected \
                    and edge.get("calls") is not None \
                    and edge["calls"] > 0 \
                    and edge.get("inclusive_ir", 0) > 0:
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
    direct_ir = sum(edge.get("inclusive_ir", 0)
                    for function_id in ids
                    for edge in functions[function_id].get("edges", []))
    if not incoming:
        state = "present_without_positive_incoming_edge"
        return {
            "target": target,
            "function_ids": ids,
            "state": state,
            "absence_reason": "no_positive_incoming_edge",
            "self_ir": self_ir,
            "direct_ir": direct_ir,
            "inclusive_ir": None,
            "calls": None,
            "incoming_edge_count": 0,
            "incoming_edges": [],
            "inline_fallback": {
                "status": "indeterminate_assembly_required",
                "candidate_callers": caller_context(parsed, label)
                if key == "claim_sector" else [],
                "value_substituted_for_target": False,
            },
            "validation": {
                "target_function_present": True,
                "positive_incoming_edge": False,
                "missing_is_not_zero_work": True,
                "assembly_required": True,
            },
        }
    return {
        "target": target,
        "function_ids": ids,
        "state": "out_of_line",
        "absence_reason": None,
        "self_ir": self_ir,
        "direct_ir": direct_ir,
        "inclusive_ir": sum(edge["inclusive_ir"] for edge in incoming),
        "calls": sum(edge["calls"] for edge in incoming),
        "incoming_edge_count": len(incoming),
        "incoming_edges": incoming,
        "inline_fallback": {
            "status": "not_required",
            "candidate_callers": [],
            "value_substituted_for_target": False,
        },
        "validation": {
            "target_function_present": True,
            "positive_incoming_edge": True,
            "missing_is_not_zero_work": True,
            "assembly_required": False,
        },
    }


def parse_dump(stage: str, job: dict[str, Any], number: int,
               path: Path) -> dict[str, Any]:
    label = relative(path)
    text = read_text(path)
    parts = [int(match.group(1)) for match in
             (PART_RE.match(line.strip()) for line in text.splitlines()) if match]
    triggers = [match.group(1) for match in
                (TRIGGER_RE.match(line.strip()) for line in text.splitlines()) if match]
    require(len(parts) == 1 and parts[0] == number,
            f"{label}: part number differs from artifact suffix")
    require(len(triggers) == 1 and triggers[0] == f"--dump-after={job['owner']}",
            f"{label}: dump trigger differs")
    events = [line.strip().split(":", 1)[1].split()
              for line in text.splitlines() if line.strip().startswith("events:")]
    positions = [line.strip().split(":", 1)[1].split()
                 for line in text.splitlines() if line.strip().startswith("positions:")]
    require(events == [["Ir"]], f"{label}: event set differs")
    require(positions == [["instr"]], f"{label}: positions differ")
    parsed = RAW.parse_raw_profile(path.resolve(), job["owner"])
    classification = classify_owner(parsed, job, label)
    targets = {
        key: attribution(parsed, key, target, label)
        for key, target in ATTRIBUTION_TARGETS.items()
    }
    return {
        "number": number,
        "role": classification["role"],
        "owner_caller": classification["owner_caller"],
        "owner_caller_id": classification["owner_caller_id"],
        "owner_calls": classification["owner_calls"],
        "owner_incoming": classification["owner_incoming"],
        "owner_self_ir": parsed["owner_self_ir"],
        "owner_direct_ir": parsed["owner_direct_ir"],
        "owner_inclusive_ir": parsed["summary_ir"],
        "summary_ir": parsed["summary_ir"],
        "runner_ancestry": classification["runner_ancestry"],
        "setup_ancestry": classification["setup_ancestry"],
        "file": parsed["file"],
        "sha256": parsed["sha256"],
        "trigger": triggers[0],
        "events": events[0],
        "positions": positions[0],
        "attribution": targets,
        "validation": {
            **classification["validation"],
            "owner_edge_matches_summary": (
                classification["owner_incoming"]["inclusive_ir"] == parsed["summary_ir"]
            ),
            "owner_self_plus_direct_matches_summary": (
                parsed["owner_self_ir"] + parsed["owner_direct_ir"] == parsed["summary_ir"]
            ),
            "target_absence_is_explicit_not_zero_work": all(
                item["validation"]["missing_is_not_zero_work"]
                for item in targets.values()
            ),
        },
    }


def numbered_paths(folder: Path, name: str) -> list[tuple[int, Path]]:
    prefix = f"{name}.callgrind."
    found = []
    for path in folder.glob(prefix + "*"):
        suffix = path.name[len(prefix):]
        if suffix.isdigit() and path.is_file() and not path.is_symlink():
            found.append((int(suffix), path))
    found.sort()
    require(found, f"{relative(folder / name)}: no numbered dumps")
    require([number for number, _path in found] == list(range(1, len(found) + 1)),
            f"{name}: numbered dumps are not contiguous")
    return found


def final_dump(path: Path, expected_part: int, owner: str) -> dict[str, Any]:
    label = relative(path)
    text = read_text(path)
    parts = [int(match.group(1)) for match in
             (PART_RE.match(line.strip()) for line in text.splitlines()) if match]
    triggers = [match.group(1) for match in
                (TRIGGER_RE.match(line.strip()) for line in text.splitlines()) if match]
    require(parts == [expected_part], f"{label}: final part differs")
    require(triggers == ["Program termination"], f"{label}: final trigger differs")
    events = [line.strip().split(":", 1)[1].split()
              for line in text.splitlines() if line.strip().startswith("events:")]
    positions = [line.strip().split(":", 1)[1].split()
                 for line in text.splitlines() if line.strip().startswith("positions:")]
    require(events == [["Ir"]] and positions == [["instr"]],
            f"{label}: final event/position set differs")
    summaries = []
    for line in text.splitlines():
        match = SUMMARY_RE.match(line.strip())
        if match:
            try:
                summaries.append(int(match.group(1).split()[0].replace(",", "")))
            except (IndexError, ValueError) as error:
                raise EvidenceError(f"{label}: malformed final summary") from error
    require(summaries and summaries[0] == 0,
            f"{label}: final process Ir is not zero")
    return {
        "file": label,
        "sha256": sha256(path),
        "part": expected_part,
        "trigger": triggers[0],
        "summary_ir": summaries[0],
        "owner": owner,
        "validation": {"program_termination_zero_ir": True},
    }


def analyze_profile(stage: str, job: dict[str, Any],
                    plan: dict[str, Any]) -> dict[str, Any]:
    folder = HERE / stage
    receipt = validate_receipt(stage, job, plan)
    report = validate_profile_report(folder, job, receipt["binary_sha256"])
    numbered = numbered_paths(folder, job["name"])
    expected_count = 5 if job["kind"] == "xls" else 6
    require(len(numbered) == expected_count,
            f"{job['name']}: numbered dump count differs")
    dumps = [parse_dump(stage, job, number, path) for number, path in numbered]
    setup = [dump for dump in dumps if dump["role"] == "setup"]
    timed = [dump for dump in dumps if dump["role"] == "timed"]
    require(len(setup) == (0 if job["kind"] == "xls" else 1),
            f"{job['name']}: setup count differs")
    require(len(timed) == plan["profile"]["samples"],
            f"{job['name']}: timed count differs")
    final = final_dump(folder / f"{job['name']}.callgrind",
                       expected_count + 1, job["owner"])
    return {
        "name": job["name"],
        "stage": stage,
        "repeat": job["repeat"],
        "group": job["group"],
        "kind": job["kind"],
        "shape": job["shape"],
        "owner": job["owner"],
        "runner": job["runner"],
        "receipt": receipt,
        "profile_result": report,
        "raw_dumps": dumps,
        "setup_dumps": setup,
        "timed_dumps": timed,
        "final_process_dump": final,
        "validation": {
            "positive_owner_edge_and_ancestry_classified": True,
            "setup_excluded_from_timed_attribution": True,
            "exactly_five_positive_timed_dumps": len(timed) == 5,
            "target_self_and_inclusive_kept_separately": True,
            "missing_or_inline_target_not_coerced_to_zero": True,
            "termination_dump_zero_ir": True,
        },
        "limitations": [
            "Callgrind Ir and call metadata are mechanism diagnostics only.",
            "Inline or absent targets require exact same-binary assembly mapping before work movement can be interpreted.",
        ],
    }


def aggregate_targets(profiles: list[dict[str, Any]]) -> dict[str, Any]:
    rows: dict[str, dict[str, Any]] = {}
    owner = {
        "inclusive_ir": 0,
        "self_ir": 0,
        "direct_ir": 0,
        "calls": 0,
        "dump_count": 0,
    }
    for profile in profiles:
        for dump in profile["timed_dumps"]:
            owner["inclusive_ir"] += dump["owner_inclusive_ir"]
            owner["self_ir"] += dump["owner_self_ir"]
            owner["direct_ir"] += dump["owner_direct_ir"]
            owner["calls"] += dump["owner_calls"]
            owner["dump_count"] += 1
            for key in ATTRIBUTION_TARGETS:
                item = dump["attribution"][key]
                row = rows.setdefault(key, {
                    "target": item["target"],
                    "self_ir": 0,
                    "direct_ir": 0,
                    "inclusive_ir": 0,
                    "calls": 0,
                    "timed_dump_count": 0,
                    "attributed_dump_count": 0,
                    "indeterminate_dump_count": 0,
                    "states": {},
                })
                state = item["state"]
                row["timed_dump_count"] += 1
                row["states"][state] = row["states"].get(state, 0) + 1
                if (item["self_ir"] is None or item["direct_ir"] is None
                        or item["inclusive_ir"] is None or item["calls"] is None):
                    row["indeterminate_dump_count"] += 1
                    continue
                row["attributed_dump_count"] += 1
                for field in ("self_ir", "direct_ir", "inclusive_ir", "calls"):
                    row[field] += item[field]
    for row in rows.values():
        if row["indeterminate_dump_count"]:
            row["state"] = "indeterminate_assembly_required"
            for field in ("self_ir", "direct_ir", "inclusive_ir", "calls"):
                row[field] = None
        else:
            row["state"] = "fully_attributed"
    return {"selected_owner": owner, "targets": rows}


def analyze_stage(stage: str, plan: dict[str, Any] | None = None) -> dict[str, Any]:
    require(stage in STAGES, f"unsupported stage {stage}")
    plan = plan or plan_data()
    profiles = [analyze_profile(stage, job, plan) for job in profile_jobs(plan)]
    require(len(profiles) == 8, f"{stage}: profile matrix is incomplete")
    require(sum(len(profile["timed_dumps"]) for profile in profiles) == 40,
            f"{stage}: expected forty timed dumps")
    require(sum(len(profile["setup_dumps"]) for profile in profiles) == 6,
            f"{stage}: expected six setup dumps")
    return {
        "schema": "ole2_physical_marker_0555_profile_analysis_v1",
        "status": "pass",
        "stage": stage,
        "scope": plan["scope"],
        "performance_claim": "diagnostic-only",
        "plan": "plan.json",
        "plan_sha256": sha256(PLAN_PATH),
        "profile_count": len(profiles),
        "timed_constructor_dump_count": 40,
        "setup_dump_count": 6,
        "profiles": profiles,
        "aggregates": aggregate_targets(profiles),
        "mechanism_gate": {
            "all_timed_owner_edges_classified": True,
            "all_cfb_setup_roles_classified_from_positive_ancestry": True,
            "selected_owner_self_and_inclusive_separate": True,
            "physical_target_absence_is_indeterminate": True,
            "claim_sector_inline_fallback_retained": True,
            "moved_work_targets_retained_separately": True,
            "no_native_or_adoption_claim": True,
        },
        "helpers": {
            "change-0554/analyze_profiles.py": sha256(OLD_ANALYZER_PATH),
            "run.py": sha256(RUN_PATH),
            "workspace-lock.json": sha256(LOCK_BINDING),
        },
        "validation": {
            "plan_receipt_binary_and_execution_bindings": True,
            "profile_reports_and_artifacts_hashed": True,
            "positive_owner_ancestry_used_for_scope": True,
            "setup_not_in_timed_aggregates": True,
            "raw_target_self_direct_inclusive_fields_separate": True,
            "missing_or_inline_values_are_not_zero": True,
            "no_native_allocation_or_adoption_claim": True,
        },
        "limitations": [
            "Callgrind Ir is a mechanism diagnostic and does not establish native latency, allocation, RSS, physical I/O, or hardware behavior.",
            "A target that is absent or inlined remains indeterminate until same-binary assembly mapping identifies the moved instructions.",
        ],
    }


def metric(baseline: int | None, candidate: int | None) -> dict[str, Any]:
    if baseline is None or candidate is None:
        return {
            "status": "indeterminate_assembly_required",
            "baseline": baseline,
            "candidate": candidate,
            "delta_ir": None,
            "delta_percent": None,
        }
    return {
        "status": "attributed",
        "baseline": baseline,
        "candidate": candidate,
        "delta_ir": candidate - baseline,
        "delta_percent": (100.0 * (candidate - baseline) / baseline)
        if baseline else None,
    }


def profile_key(profile: dict[str, Any]) -> tuple[int, str, str | None]:
    return profile["repeat"], profile["group"], profile["shape"]


def aggregate_profile_target(profile: dict[str, Any], key: str) -> dict[str, Any]:
    values = [dump["attribution"][key] for dump in profile["timed_dumps"]]
    fields = ("self_ir", "direct_ir", "inclusive_ir", "calls")
    if any(any(item[field] is None for field in fields) for item in values):
        return {
            "state": "indeterminate_assembly_required",
            **{field: None for field in fields},
            "timed_dump_count": len(values),
            "attributed_dump_count": sum(
                all(item[field] is not None for field in fields) for item in values
            ),
        }
    return {
        "state": "fully_attributed",
        **{field: sum(item[field] for item in values) for field in fields},
        "timed_dump_count": len(values),
        "attributed_dump_count": len(values),
    }


def compare_stages(baseline: dict[str, Any],
                   candidate: dict[str, Any]) -> dict[str, Any]:
    left = {profile_key(item): item for item in baseline["profiles"]}
    right = {profile_key(item): item for item in candidate["profiles"]}
    require(set(left) == set(right), "profile matrices differ")
    rows = []
    for key in sorted(left):
        bp = left[key]
        cp = right[key]
        target_rows = {}
        for target in ATTRIBUTION_TARGETS:
            b = aggregate_profile_target(bp, target)
            c = aggregate_profile_target(cp, target)
            target_rows[target] = {
                "baseline_state": b["state"],
                "candidate_state": c["state"],
                "self_ir": metric(b["self_ir"], c["self_ir"]),
                "direct_ir": metric(b["direct_ir"], c["direct_ir"]),
                "inclusive_ir": metric(b["inclusive_ir"], c["inclusive_ir"]),
                "calls": metric(b["calls"], c["calls"]),
                "baseline_attributed_dump_count": b["attributed_dump_count"],
                "candidate_attributed_dump_count": c["attributed_dump_count"],
            }
        b_owner = sum(dump["owner_inclusive_ir"] for dump in bp["timed_dumps"])
        c_owner = sum(dump["owner_inclusive_ir"] for dump in cp["timed_dumps"])
        rows.append({
            "repeat": key[0],
            "group": key[1],
            "shape": key[2],
            "baseline_profile": bp["name"],
            "candidate_profile": cp["name"],
            "owner_inclusive_ir": metric(b_owner, c_owner),
            "targets": target_rows,
            "claim_sector_inline_context": {
                "baseline": [dump["attribution"]["claim_sector"]["inline_fallback"]
                             for dump in bp["timed_dumps"]
                             if dump["attribution"]["claim_sector"]["state"] != "out_of_line"],
                "candidate": [dump["attribution"]["claim_sector"]["inline_fallback"]
                              for dump in cp["timed_dumps"]
                              if dump["attribution"]["claim_sector"]["state"] != "out_of_line"],
                "target_cost_not_replaced_by_caller_cost": True,
            },
        })

    row_by_key = {(row["repeat"], row["group"], row["shape"]): row for row in rows}
    xls_owner_rows = [row_by_key[(repeat, "xls-owned", None)] for repeat in (1, 2)]
    # The selected physical-reconciliation leaf is measured in the owned XLS
    # constructor lane that exposed the residual 17.6% hotspot.  CFB shape
    # rows remain in the report as controls and moved-work diagnostics; they
    # are not silently substituted for this targeted mechanism gate.
    physical_rows = [row for row in rows if row["group"] == "xls-owned"]
    physical_metrics = [
        row["targets"]["physical_reconciliation"]["self_ir"]
        for row in physical_rows
    ]
    physical_present = all(
        row["targets"]["physical_reconciliation"]["self_ir"]["status"] == "attributed"
        for row in physical_rows
    )
    physical_decreases = (
        physical_present and all(
            row["targets"]["physical_reconciliation"]["self_ir"]["delta_ir"] < 0
            for row in physical_rows
        )
    )
    claim_states = [
        dump["attribution"]["claim_sector"]["state"]
        for profile in [*left.values(), *right.values()]
        for dump in profile["timed_dumps"]
    ]
    return {
        "rows": rows,
        "moved_work": {
            "load_fat": [row["targets"]["load_fat"] for row in rows],
            "claim_sector_or_inline_caller": [
                {
                    "repeat": row["repeat"],
                    "group": row["group"],
                    "shape": row["shape"],
                    "target": row["targets"]["claim_sector"],
                    "inline_context": row["claim_sector_inline_context"],
                }
                for row in rows
            ],
            "validate_stream_allocations": [
                row["targets"]["validate_stream_allocations"] for row in rows
            ],
            "collect_exact": [row["targets"]["collect_exact"] for row in rows],
            "physical_reconciliation": [
                row["targets"]["physical_reconciliation"] for row in rows
            ],
            "interpretation": (
                "These rows retain where measured Ir is attributed; no row is "
                "interpreted as eliminated work. Inline or absent targets require "
                "same-binary assembly mapping."
            ),
        },
        "mechanism_gate": {
            "xls_owner_inclusive_ir_decreases_each_repeat": all(
                row["owner_inclusive_ir"]["delta_ir"] < 0 for row in xls_owner_rows
            ),
            "physical_reconciliation_self_ir_decreases_each_repeat": bool(physical_decreases),
            "physical_reconciliation_self_ir_status": (
                "proven" if physical_present else "indeterminate_assembly_required"
            ),
            "physical_reconciliation_present_for_all_xls_rows": physical_present,
            "all_target_rows_kept_separate": True,
            "claim_sector_inline_or_out_of_line_explicit": all(
                state in {
                    "out_of_line", "inlined_or_absent",
                    "present_without_positive_incoming_edge",
                }
                for state in claim_states
            ),
            "moved_work_retained_without_elimination_claim": True,
            "no_native_or_adoption_claim": True,
        },
        "validation": {
            "profile_matrix_matches": True,
            "owner_self_and_inclusive_compared_separately": True,
            "all_requested_targets_compared_separately": True,
            "inline_or_absent_not_coerced_to_zero": True,
            "physical_leaf_requires_assembly_when_missing": True,
        },
        "interpretation": (
            "The XLS owner direction and CFB physical leaf direction are "
            "mechanism gates only.  A missing or inlined physical leaf is "
            "indeterminate and cannot satisfy the leaf direction gate until "
            "assembly mapping is available."
        ),
    }


def compare(plan: dict[str, Any] | None = None) -> dict[str, Any]:
    plan = plan or plan_data()
    baseline = analyze_stage("baseline", plan)
    candidate = analyze_stage("candidate", plan)
    comparison = compare_stages(baseline, candidate)
    return {
        "schema": "ole2_physical_marker_0555_profile_comparison_v1",
        "status": "pass",
        "scope": plan["scope"],
        "performance_claim": "diagnostic-only",
        "plan": "plan.json",
        "plan_sha256": sha256(PLAN_PATH),
        "stage_selection": ["baseline", "candidate"],
        "stages": {"baseline": baseline, "candidate": candidate},
        "comparison": comparison,
        "validation": {
            "both_stages_valid": True,
            "matched_owner_edges_compared": True,
            "physical_marker_targets_retained": True,
            "inline_or_absent_not_zero_work": True,
            "no_native_or_adoption_claim": True,
        },
        "limitations": [
            "Callgrind Ir and positive call metadata are mechanism diagnostics only.",
            "The physical leaf gate is indeterminate when either stage lacks a positive emitted target; same-binary assembly mapping is required.",
        ],
    }


def write_report(document: dict[str, Any], output: Path) -> None:
    canonical = {
        HERE / "baseline" / "profile-analysis.json",
        HERE / "candidate" / "profile-analysis.json",
        HERE / "profile-comparison.json",
    }
    resolved = output.resolve()
    require(resolved in {path.resolve() for path in canonical},
            f"refusing non-canonical profile output {output}")
    data = (json.dumps(document, indent=2, sort_keys=True) + "\n").encode("utf-8")
    if output.exists():
        require(output.is_file() and not output.is_symlink(),
                f"refusing non-regular output {output}")
        require(output.read_bytes() == data,
                f"refusing to overwrite non-identical output {output}")
        return
    require(output.parent.is_dir(), f"canonical output directory is missing: {output.parent}")
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
    output = args.output_option or args.output or (
        HERE / "profile-comparison.json" if args.compare else
        HERE / args.stage / "profile-analysis.json"
    )
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
