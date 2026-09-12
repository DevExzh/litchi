#!/usr/bin/env python3
"""Validate the 0523 CFB/OLE2 constructor Callgrind profiles.

The profile lane is a baseline attribution run.  It contains one XLS
source-open child for each repeat and one CFB child for each repeat/shape.
Each child runs five measured constructor calls with ``--zero-before`` and
``--dump-after`` bound to the selected constructor.  CFB corpus generation
also opens the generated file once before the timed loop, so its first
numbered dump is retained as setup evidence and is excluded from constructor
attribution.  XLS has no selected-constructor setup call.

This analyzer proves scope from the positive raw incoming edge into the
selected constructor.  It does not infer a timed dump from its ordinal, and
it does not compare a candidate or make a speedup claim.  Raw Callgrind
metadata is a mechanism diagnostic: with collection disabled outside the
toggle, child-call metadata can retain context from collection-off work.  The
selected incoming edge, owner total, and native output identity are therefore
the authoritative scope checks.
"""

from __future__ import annotations

import argparse
import datetime
import importlib.util
import json
import os
from pathlib import Path
import re
import subprocess
import sys
from typing import Any


HERE = Path(__file__).resolve().parent
PLAN_PATH = HERE / "plan.json"
RUN_PATH = HERE / "run.py"
RETAINED_DIR = HERE.parent / "change-0519"
ROOT_ANALYZER = HERE / "analyze.py"

XLS_CASE = "xls_owned_source_open_one_cell"
CFB_CASE = "cfb_open"
XLS_RUNNER = "litchi_perf_baseline::run_xls_owned_source_case"
CFB_RUNNER = "litchi_perf_baseline::run_cfb_open"
CFB_SETUP = "litchi_perf_baseline::build_cfb_corpus"
CFB_SETUP_INLINE_CALLER = "litchi_perf_baseline::run::{{closure}}"
XLS_OWNER_WRAPPER = (
    "litchi_xls::workbook::source::SourceBackedWorkbook::from_read_at"
)
MAX_RUNNER_ANCESTRY_DEPTH = 8

CHAIN_COLLECTOR = "litchi_cfb::file::SectorChainScratch::collect_exact"
CLAIM_SECTOR = "litchi_cfb::file::OleFile<R>::claim_sector"
VALIDATE_STREAM_ALLOCATIONS = (
    "litchi_cfb::file::OleFile<R>::validate_stream_allocations"
)
PHYSICAL_RECONCILIATION = (
    "litchi_cfb::file::OleFile<R>::validate_physical_sector_layout"
)
LOAD_FAT = "litchi_cfb::file::OleFile<R>::load_fat"

ATTRIBUTION_TARGETS = {
    "chain_collector": CHAIN_COLLECTOR,
    "claim_sector": CLAIM_SECTOR,
    "validate_stream_allocations": VALIDATE_STREAM_ALLOCATIONS,
    "physical_reconciliation": PHYSICAL_RECONCILIATION,
    "load_fat": LOAD_FAT,
}

FUNCTION_RE = re.compile(r"^(fn|cfn)=\((\d+)\)(?:\s+(.*))?$")
CALLS_RE = re.compile(r"^calls=([\d,]+)")
SUMMARY_RE = re.compile(r"^summary:\s*(.*)$")
PART_RE = re.compile(r"^part:\s*(\d+)$")
TRIGGER_RE = re.compile(r"^desc:\s+Trigger:\s+(.*)$")
SHA256_RE = re.compile(r"^[0-9a-f]{64}$")
WARNING_RE = re.compile(
    r"^(?:warning|overflow|error|failed|notice)(?:\s|:|$)",
    re.IGNORECASE,
)
PERL_ENV = {"PERL_HASH_SEED": "0", "PERL_PERTURB_KEYS": "0"}


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


def check_hash(value: Any, label: str) -> str:
    require(isinstance(value, str) and SHA256_RE.fullmatch(value) is not None,
            f"{label}: expected lowercase SHA-256")
    return value


def read_json(path: Path) -> Any:
    try:
        with path.open(encoding="utf-8") as stream:
            return json.load(stream)
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise EvidenceError(f"cannot read JSON {path}: {error}") from error


def read_text(path: Path) -> str:
    try:
        return path.read_text(encoding="utf-8", errors="replace")
    except OSError as error:
        raise EvidenceError(f"cannot read {path}: {error}") from error


def relative(path: Path) -> str:
    try:
        return str(path.relative_to(HERE))
    except ValueError as error:
        raise EvidenceError(f"path is outside evidence directory: {path}") from error


def load_helpers() -> tuple[Any, Any]:
    """Load immutable 0519 raw-edge helpers without changing prior evidence."""

    raw_path = RETAINED_DIR / "analyze_profiles.py"
    edges_path = RETAINED_DIR / "compare_profile_lanes.py"
    require(raw_path.is_file(), f"missing retained raw helper: {raw_path}")
    require(edges_path.is_file(), f"missing retained edge helper: {edges_path}")

    raw_spec = importlib.util.spec_from_file_location(
        "litchi_0523_retained_raw_profiles", raw_path
    )
    require(raw_spec is not None and raw_spec.loader is not None,
            f"cannot load retained raw helper: {raw_path}")
    raw = importlib.util.module_from_spec(raw_spec)
    raw_spec.loader.exec_module(raw)

    # The retained edge module imports its parser by this historical module
    # name.  Restore any caller-owned module after loading it.
    previous = sys.modules.get("analyze_profiles")
    sys.modules["analyze_profiles"] = raw
    try:
        edge_spec = importlib.util.spec_from_file_location(
            "litchi_0523_retained_profile_edges", edges_path
        )
        require(edge_spec is not None and edge_spec.loader is not None,
                f"cannot load retained edge helper: {edges_path}")
        edges = importlib.util.module_from_spec(edge_spec)
        edge_spec.loader.exec_module(edges)
    finally:
        if previous is None:
            sys.modules.pop("analyze_profiles", None)
        else:
            sys.modules["analyze_profiles"] = previous
    return raw, edges


RAW, EDGES = load_helpers()


def load_numeric_helper() -> Any:
    """Load the companion numerical verifier only when evidence is analyzed."""

    require(ROOT_ANALYZER.is_file(), f"missing companion numerical analyzer: {ROOT_ANALYZER}")
    old_path = list(sys.path)
    old_run = sys.modules.get("run")
    # analyze.py imports run.py by its historical script name.  Install the
    # current 0523 driver explicitly so a caller that already imported a
    # different campaign's ``run`` module cannot redirect the helper.
    if str(HERE) not in sys.path:
        sys.path.insert(0, str(HERE))
    try:
        run_spec = importlib.util.spec_from_file_location(
            "litchi_0523_profile_run", RUN_PATH
        )
        require(run_spec is not None and run_spec.loader is not None,
                f"cannot load current run driver: {RUN_PATH}")
        current_run = importlib.util.module_from_spec(run_spec)
        sys.modules["run"] = current_run
        run_spec.loader.exec_module(current_run)
        spec = importlib.util.spec_from_file_location(
            "litchi_0523_numeric_profile_helper", ROOT_ANALYZER
        )
        require(spec is not None and spec.loader is not None,
                f"cannot load companion numerical analyzer: {ROOT_ANALYZER}")
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        return module
    except AssertionError as error:
        raise EvidenceError(f"companion numerical analyzer rejected its setup: {error}") from error
    finally:
        sys.path[:] = old_path
        if old_run is None:
            sys.modules.pop("run", None)
        else:
            sys.modules["run"] = old_run


def plan_data() -> dict[str, Any]:
    plan = read_json(PLAN_PATH)
    require(isinstance(plan, dict), "plan is not an object")
    profile = plan.get("profile")
    require(isinstance(profile, dict), "plan profile is not an object")
    require(profile.get("xls_owner") == (
        "litchi_xls::workbook::source::SourceBackedWorkbook::from_read_at_with_limits"
    ), "plan XLS profile owner differs")
    require(profile.get("cfb_owner") == "litchi_cfb::file::OleFile<R>::open",
            "plan CFB profile owner differs")
    require(profile.get("repeats") == 2, "profile repeat count differs")
    require(profile.get("warmup") == 0, "profile warmup differs")
    require(profile.get("samples") == 5, "profile sample count differs")
    require(profile.get("jobs") == [
        "xls-owned", "cfb-tiny", "cfb-many-small", "cfb-few-large"
    ], "profile job matrix differs")
    require(plan.get("cpu") == 2, "profile CPU differs from the pinned plan")
    owned = plan.get("owned_paths")
    require(owned == ["/tmp/litchi-goal-0523", "/home/zhuhe/litchi-goal-0523-target"],
            "owned temporary paths differ from the plan")
    groups = plan.get("groups")
    require(isinstance(groups, dict), "plan groups are missing")
    require(groups.get("xls", {}).get("cases") and
            XLS_CASE in groups["xls"]["cases"],
            "XLS native group must retain the profiled owned-source case")
    require(groups.get("cfb", {}).get("cases") == [CFB_CASE],
            "CFB native group differs from cfb_open")
    require(groups.get("cfb", {}).get("shapes") == ["tiny", "many-small", "few-large"],
            "CFB profile shape matrix differs")
    return plan


def profile_jobs(plan: dict[str, Any]) -> list[dict[str, Any]]:
    profile = plan["profile"]
    jobs: list[dict[str, Any]] = []
    for repeat in range(1, profile["repeats"] + 1):
        jobs.append({
            "name": f"profile-r{repeat}-xls-owned",
            "repeat": repeat,
            "group": "xls-owned",
            "shape": None,
            "owner": profile["xls_owner"],
            "runner": XLS_RUNNER,
            "setup": None,
            "setup_callers": [],
            "samples": profile["samples"],
            "warmup": profile["warmup"],
            "selection": {"cases": [XLS_CASE]},
        })
        for shape in plan["groups"]["cfb"]["shapes"]:
            jobs.append({
                "name": f"profile-r{repeat}-cfb-{shape}",
                "repeat": repeat,
                "group": f"cfb-{shape}",
                "shape": shape,
                "owner": profile["cfb_owner"],
                "runner": CFB_RUNNER,
                "setup": CFB_SETUP,
                "setup_callers": [CFB_SETUP, CFB_SETUP_INLINE_CALLER],
                "samples": profile["samples"],
                "warmup": profile["warmup"],
                "selection": {
                    "cases": [CFB_CASE],
                    "shapes": [shape],
                    "payload": "incompressible",
                },
            })
    return jobs


def native_jobs(plan: dict[str, Any]) -> list[dict[str, Any]]:
    native = plan["native"]
    return [
        {
            "name": f"native-r{repeat}-{group}",
            "repeat": repeat,
            "group": group,
            "selection": plan["groups"][group],
            "samples": native["samples"],
            "warmup": native["warmup"],
        }
        for repeat in range(1, native["repeats"] + 1)
        for group in ("xls", "cfb")
    ]


def option(command: list[str], name: str) -> str:
    values: list[str] = []
    for index, item in enumerate(command):
        if item == name:
            require(index + 1 < len(command), f"profile command omits value for {name}")
            values.append(command[index + 1])
        elif item.startswith(name + "="):
            values.append(item.split("=", 1)[1])
    require(len(values) == 1, f"profile command omits or repeats {name}")
    return values[0]


def validate_receipt_artifacts(stage_dir: Path, receipt: dict[str, Any],
                               expected: set[str], label: str) -> None:
    artifacts = receipt.get("artifacts")
    require(isinstance(artifacts, dict), f"{label}: artifacts are not an object")
    actual = set(artifacts)
    require(actual == expected,
            f"{label}: artifact set differs; missing={sorted(expected - actual)}, "
            f"extra={sorted(actual - expected)}")
    for filename, digest in artifacts.items():
        require(isinstance(filename, str) and Path(filename).name == filename,
                f"{label}: artifact is not a stage-local basename: {filename!r}")
        check_hash(digest, f"{label}/{filename} receipt hash")
        path = stage_dir / filename
        require(path.is_file() and not path.is_symlink(),
                f"{label}: artifact is missing: {filename}")
        require(sha256(path) == digest, f"{label}: artifact hash differs: {filename}")


def validate_cleanup(plan: dict[str, Any]) -> dict[str, Any]:
    """Require explicit custody if the retained binary has been cleaned."""

    path = HERE / "cleanup.json"
    require(path.is_file(), f"cleanup record is required after owned paths are removed: {path}")
    cleanup = read_json(path)
    require(isinstance(cleanup, dict), "cleanup record is not an object")
    require(cleanup.get("owned_paths_absent") is True,
            "cleanup record does not prove owned paths are absent")
    removed = cleanup.get("removed", cleanup.get("removed_paths"))
    require(isinstance(removed, list), "cleanup record has no removed-path list")
    require(set(plan["owned_paths"]).issubset(removed),
            "cleanup record does not name every planned owned path")
    require(not cleanup.get("accessible_process_references"),
            "cleanup record retains accessible process references")
    return {
        "file": relative(path),
        "sha256": sha256(path),
        "owned_paths_absent": True,
        "removed": removed,
        "accessible_process_references": cleanup.get("accessible_process_references", []),
        "python_cache_absent": cleanup.get("python_cache_absent"),
    }


def validate_build(plan: dict[str, Any]) -> dict[str, Any]:
    stage_dir = HERE / "baseline"
    source_manifest = stage_dir / "source-manifest.json"
    receipt_path = stage_dir / "build-normal.receipt.json"
    binary_path = stage_dir / "binary-normal.json"
    for path, label in ((source_manifest, "source manifest"),
                        (receipt_path, "normal build receipt"),
                        (binary_path, "normal binary identity")):
        require(path.is_file(), f"baseline {label} is missing: {path}")
    manifest_sha = sha256(source_manifest)
    plan_sha = sha256(PLAN_PATH)
    run_sha = sha256(RUN_PATH)
    receipt = read_json(receipt_path)
    require(receipt.get("exit_code") == 0, "baseline normal build failed")
    require(receipt.get("binary_sha256") is None,
            "normal build receipt unexpectedly carries a runtime binary hash")
    require(receipt.get("source_manifest_sha256") == manifest_sha,
            "normal build/source manifest binding differs")
    require(receipt.get("plan_sha256") == plan_sha, "normal build/plan binding differs")
    require(receipt.get("script_sha256") == run_sha, "normal build/run binding differs")
    require(receipt.get("seconds", 0) > 0, "normal build has no positive duration")
    try:
        start = datetime.datetime.fromisoformat(receipt["start_utc"])
        end = datetime.datetime.fromisoformat(receipt["end_utc"])
    except (KeyError, TypeError, ValueError) as error:
        raise EvidenceError("normal build receipt timestamps are malformed") from error
    require(start < end, "normal build receipt interval is not positive")
    environment = receipt.get("environment")
    require(isinstance(environment, dict),
            "normal build receipt environment is not an object")
    require(all(value is None for value in environment.values()),
            "normal build inherited an unrecorded instrumentation environment")
    validate_receipt_artifacts(
        stage_dir, receipt, {"build-normal.stdout", "build-normal.stderr"},
        "baseline/build-normal"
    )
    expected_command = [
        "env", "CARGO_BUILD_JOBS=2", "CARGO_INCREMENTAL=0", "cargo", "build",
        "--release", "--locked", "--manifest-path", "tools/perf-baseline/Cargo.toml",
        "--bin", "litchi-perf-baseline", "--target-dir", plan["owned_paths"][1],
    ]
    require(receipt.get("command") == expected_command,
            "normal build command differs from the frozen driver")

    binary = read_json(binary_path)
    require(isinstance(binary, dict), "normal binary identity is not an object")
    binary_sha = check_hash(binary.get("sha256"), "normal binary identity")
    require(binary.get("source_manifest_sha256") == manifest_sha,
            "normal binary/source manifest binding differs")
    require(binary.get("build_receipt_sha256") == sha256(receipt_path),
            "normal binary/build receipt binding differs")
    binary_file = Path(binary.get("path", ""))
    expected_binary = Path(plan["owned_paths"][0]) / "normal"
    require(binary_file == expected_binary,
            f"normal binary path differs: {binary_file} != {expected_binary}")
    if binary_file.is_file() and not binary_file.is_symlink():
        require(sha256(binary_file) == binary_sha,
                "available normal binary hash differs from identity")
        cleanup = None
    else:
        cleanup = validate_cleanup(plan)
    require(binary.get("bytes", 0) > 0, "normal binary identity has no positive size")
    return {
        "source_manifest": relative(source_manifest),
        "source_manifest_sha256": manifest_sha,
        "build_receipt": relative(receipt_path),
        "build_receipt_sha256": sha256(receipt_path),
        "binary_identity": relative(binary_path),
        "binary_sha256": binary_sha,
        "binary_bytes": binary["bytes"],
        "binary_path": str(binary_file),
        "plan_sha256": plan_sha,
        "run_script_sha256": run_sha,
        "binary_available": binary_file.is_file() and not binary_file.is_symlink(),
        "cleanup": cleanup,
    }


def parse_cost(line: str, event_index: int) -> int | None:
    try:
        return RAW._cost(line, event_index)
    except (AttributeError, TypeError, ValueError):
        fields = line.split()
        if len(fields) <= event_index + 1:
            return None
        position = fields[0]
        if not position or (position[0] not in "+-*" and not position[0].isdigit()):
            return None
        try:
            return int(fields[event_index + 1])
        except ValueError:
            return None


def parse_raw_profile(path: Path, selected: str) -> dict[str, Any]:
    """Parse function sections and positive raw edges from one dump."""

    text = read_text(path)
    events: list[str] = []
    summary_values: list[int] | None = None
    names: dict[int, str] = {}
    functions: dict[int, dict[str, Any]] = {}
    current: int | None = None
    pending_callee: int | None = None
    pending_calls: int | None = None
    pending_line: int | None = None

    for line_number, raw_line in enumerate(text.splitlines(), 1):
        line = raw_line.strip()
        if not line:
            continue
        if line.startswith("events:"):
            events = line.split(":", 1)[1].split()
            continue
        summary_match = SUMMARY_RE.match(line)
        if summary_match:
            values: list[int] = []
            for token in summary_match.group(1).split():
                try:
                    values.append(int(token.replace(",", "")))
                except ValueError:
                    break
            summary_values = values
            continue
        function_match = FUNCTION_RE.match(line)
        if function_match:
            kind, function_text_id, function_name = function_match.groups()
            function_id = int(function_text_id)
            if function_name:
                names[function_id] = function_name
            function = functions.setdefault(
                function_id, {"name": "", "self_ir": 0, "edges": []}
            )
            if function_name:
                function["name"] = function_name
            if kind == "fn":
                current = function_id
                pending_callee = pending_calls = pending_line = None
            else:
                require(current is not None,
                        f"{relative(path)}:{line_number}: cfn outside fn")
                pending_callee = function_id
                pending_calls = None
                pending_line = line_number
            continue
        calls_match = CALLS_RE.match(line)
        if calls_match:
            require(current is not None and pending_callee is not None,
                    f"{relative(path)}:{line_number}: calls outside cfn")
            require(pending_calls is None,
                    f"{relative(path)}:{line_number}: repeated calls record")
            pending_calls = int(calls_match.group(1).replace(",", ""))
            continue
        if line.startswith((
            "creator:", "cmd:", "pid:", "version:", "positions:", "part:",
            "ob=", "cob=", "fl=", "cfi=", "fi=", "fe=", "totals:",
        )):
            continue
        if current is None or "Ir" not in events:
            continue
        cost = parse_cost(line, events.index("Ir"))
        if cost is None:
            continue
        if pending_callee is not None:
            functions[current]["edges"].append({
                "callee_id": pending_callee,
                "calls": pending_calls,
                "inclusive_ir": cost,
                "line": pending_line,
            })
        else:
            functions[current]["self_ir"] += cost
        pending_callee = pending_calls = pending_line = None

    require(events == ["Ir"], f"{relative(path)}: Callgrind event set is not exactly Ir")
    require(summary_values is not None and summary_values,
            f"{relative(path)}: Callgrind summary is missing")
    summary_ir = summary_values[0]
    for function_id, function in functions.items():
        if not function["name"]:
            function["name"] = names.get(function_id, "")
    selected_ids = [
        function_id for function_id, function in functions.items()
        if selected in function["name"]
    ]
    require(selected_ids, f"{relative(path)}: selected function {selected!r} is absent")

    incoming = []
    selected_set = set(selected_ids)
    for parent_id, function in functions.items():
        for edge in function["edges"]:
            if edge["callee_id"] in selected_set and edge["calls"] is not None \
                    and edge["calls"] > 0 and edge["inclusive_ir"] > 0:
                incoming.append({
                    "caller_id": parent_id,
                    "caller": function["name"] or names.get(parent_id, ""),
                    "callee_id": edge["callee_id"],
                    "calls": edge["calls"],
                    "inclusive_ir": edge["inclusive_ir"],
                    "line": edge["line"],
                })
    # Preserve every direct cost in the owner equation.  Callgrind normally
    # emits non-negative Ir costs, but retaining a zero or negative record is
    # safer than silently dropping it before checking the raw accounting.
    direct_ir = sum(
        edge["inclusive_ir"]
        for function_id in selected_ids
        for edge in functions[function_id]["edges"]
    )
    self_ir = sum(functions[function_id]["self_ir"] for function_id in selected_ids)
    require(len(incoming) == 1,
            f"{relative(path)}: selected function has {len(incoming)} positive incoming edges")
    owner_edge = incoming[0]
    require(owner_edge["inclusive_ir"] == summary_ir,
            f"{relative(path)}: owner incoming Ir does not equal dump summary")
    require(owner_edge["inclusive_ir"] == self_ir + direct_ir,
            f"{relative(path)}: owner self plus direct Ir does not equal incoming Ir")

    # Cross-check the retained raw edge parser.  This keeps the 0523 function
    # attribution parser independent while proving that its scope edge agrees
    # with the immutable helper used by earlier evidence.
    retained = EDGES.raw_incoming_call_summary(path, selected)
    require(retained["positive_edge_count"] == 1,
            f"{relative(path)}: retained parser found a different incoming-edge count")
    retained_edge = retained["edges"][0]
    require(
        (retained_edge["caller"], retained_edge["calls"], retained_edge["inclusive_ir"])
        == (owner_edge["caller"], owner_edge["calls"], owner_edge["inclusive_ir"]),
        f"{relative(path)}: retained/raw incoming-edge parsers disagree",
    )
    return {
        "file": relative(path),
        "sha256": sha256(path),
        "summary_ir": summary_ir,
        "selected_function": selected,
        "selected_function_ids": selected_ids,
        "owner_incoming": owner_edge,
        "owner_self_ir": self_ir,
        "owner_direct_ir": direct_ir,
        "functions": functions,
        "incoming_edges": incoming,
        "warnings": [line for line in text.splitlines() if WARNING_RE.search(line)],
        "validation": {
            "event_set_exactly_ir": True,
            "one_positive_owner_incoming_edge": True,
            "owner_edge_matches_summary": True,
            "owner_self_plus_direct_matches_incoming": True,
            "retained_raw_edge_parser_agrees": True,
        },
    }


def part_number(text: str, label: str) -> int:
    values = [int(match.group(1)) for match in
              (PART_RE.match(line.strip()) for line in text.splitlines()) if match]
    require(len(values) == 1, f"{label}: expected one part record, got {values}")
    return values[0]


def trigger(text: str, label: str) -> str:
    values = [match.group(1) for match in
              (TRIGGER_RE.match(line.strip()) for line in text.splitlines()) if match]
    require(len(values) == 1, f"{label}: expected one Trigger description, got {values}")
    return values[0]


def caller_matches(actual: str, expected: str) -> bool:
    return actual == expected or actual.endswith("::" + expected.rsplit("::", 1)[-1])


def positive_ancestor_path(parsed: dict[str, Any], target_id: int,
                           ancestor: str) -> dict[str, Any] | None:
    """Find a bounded positive-call path from ``ancestor`` to ``target_id``.

    The selected constructor's immediate caller can be a library convenience
    wrapper.  A direct ``runner -> constructor`` edge is therefore not a
    requirement: the owner edge is checked separately, and this search proves
    that its caller is reachable from the benchmark runner through positive
    raw call edges.  The short bound prevents unrelated recursive callgraph
    paths from turning into scope evidence.
    """

    functions = parsed["functions"]
    starts = sorted(
        function_id for function_id, function in functions.items()
        if caller_matches(function["name"], ancestor)
    )
    if not starts:
        return None

    queue: list[tuple[int, list[int], list[dict[str, Any]]]] = [
        (function_id, [function_id], []) for function_id in starts
    ]
    visited = set(starts)
    while queue:
        current_id, path_ids, path_edges = queue.pop(0)
        if current_id == target_id:
            return {
                "ancestor": ancestor,
                "function_ids": path_ids,
                "functions": [functions.get(function_id, {}).get("name", "")
                              for function_id in path_ids],
                "edges": path_edges,
                "depth": len(path_edges),
                "max_depth": MAX_RUNNER_ANCESTRY_DEPTH,
            }
        if len(path_edges) >= MAX_RUNNER_ANCESTRY_DEPTH:
            continue
        current = functions.get(current_id)
        if current is None:
            continue
        edges = sorted(
            (
                edge for edge in current["edges"]
                if edge["calls"] is not None
                # With collection disabled outside the exact owner toggle,
                # Valgrind can retain a positive-cost ancestry edge with a
                # zero call count.  It remains useful as context evidence;
                # only the selected owner's own edge requires calls == 1.
                and edge["calls"] >= 0
                and edge["inclusive_ir"] > 0
            ),
            key=lambda edge: (
                functions.get(edge["callee_id"], {}).get("name", ""),
                edge["callee_id"],
                edge["line"] or 0,
            ),
        )
        for edge in edges:
            callee_id = edge["callee_id"]
            if callee_id in visited:
                continue
            visited.add(callee_id)
            queue.append((
                callee_id,
                [*path_ids, callee_id],
                [*path_edges, {
                    "caller_id": current_id,
                    "caller": functions.get(current_id, {}).get("name", ""),
                    "callee_id": callee_id,
                    "callee": functions.get(callee_id, {}).get("name", ""),
                    "calls": edge["calls"],
                    "inclusive_ir": edge["inclusive_ir"],
                    "line": edge["line"],
                }],
            ))
    return None


def direct_context_path(parsed: dict[str, Any], function_id: int,
                        ancestor: str) -> dict[str, Any]:
    """Represent a known immediate caller as a zero-edge context path."""

    name = parsed["functions"].get(function_id, {}).get("name", "")
    return {
        "ancestor": ancestor,
        "function_ids": [function_id],
        "functions": [name],
        "edges": [],
        "depth": 0,
        "max_depth": MAX_RUNNER_ANCESTRY_DEPTH,
        "direct_owner_caller": True,
    }


def validate_numbered_dump(path: Path, number: int, job: dict[str, Any],
                           plan: dict[str, Any]) -> dict[str, Any]:
    text = read_text(path)
    label = relative(path)
    require(part_number(text, label) == number,
            f"{label}: part number does not match suffix .{number}")
    require(trigger(text, label) == f"--dump-after={job['owner']}",
            f"{label}: Trigger does not bind the selected owner")
    parsed = parse_raw_profile(path, job["owner"])
    caller = parsed["owner_incoming"]["caller"]
    if job["shape"] is None and not caller_matches(caller, job["runner"]):
        require(caller_matches(caller, XLS_OWNER_WRAPPER),
                f"{label}: XLS owner immediate caller {caller!r} is not the "
                f"expected wrapper {XLS_OWNER_WRAPPER!r}")
    owner_caller_id = parsed["owner_incoming"]["caller_id"]
    setup_callers = job.get("setup_callers", [])
    direct_runner = caller_matches(caller, job["runner"])
    direct_setup = [expected for expected in setup_callers
                    if caller_matches(caller, expected)]
    if direct_runner and direct_setup:
        raise EvidenceError(
            f"{label}: positive owner caller {caller!r} matches both timed and setup callers"
        )
    if direct_runner:
        runner_path = direct_context_path(parsed, owner_caller_id, job["runner"])
        setup_path = None
        role = "timed"
        require(parsed["owner_incoming"]["calls"] == 1,
                f"{label}: timed owner edge contains multiple constructor calls")
    elif direct_setup:
        runner_path = None
        setup_path = direct_context_path(parsed, owner_caller_id, direct_setup[0])
        role = "setup"
        require(number == 1,
                f"{label}: CFB setup edge appeared after the first numbered dump")
        require(parsed["owner_incoming"]["calls"] == 1,
                f"{label}: setup owner edge contains multiple constructor calls")
    else:
        runner_path = positive_ancestor_path(parsed, owner_caller_id, job["runner"])
        setup_paths = [
            positive_ancestor_path(parsed, owner_caller_id, expected)
            for expected in setup_callers
        ]
        setup_path = next((candidate for candidate in setup_paths if candidate is not None), None)
        if runner_path is not None and setup_path is not None:
            raise EvidenceError(
                f"{label}: positive owner caller {caller!r} is reachable from both "
                f"the timed runner and setup caller"
            )
        if runner_path is not None:
            role = "timed"
            require(parsed["owner_incoming"]["calls"] == 1,
                    f"{label}: timed owner edge contains multiple constructor calls")
        elif setup_path is not None:
            role = "setup"
            require(number == 1,
                    f"{label}: CFB setup edge appeared after the first numbered dump")
            require(parsed["owner_incoming"]["calls"] == 1,
                    f"{label}: setup owner edge contains multiple constructor calls")
        else:
            raise EvidenceError(
                f"{label}: positive owner caller {caller!r} is neither the timed "
                f"runner {job['runner']!r} nor reachable from an allowed setup caller "
                f"{setup_callers!r}"
            )
    return {
        "number": number,
        "role": role,
        "owner_caller": caller,
        "owner_calls": parsed["owner_incoming"]["calls"],
        "summary_ir": parsed["summary_ir"],
        "owner_incoming": parsed["owner_incoming"],
        "owner_self_ir": parsed["owner_self_ir"],
        "owner_direct_ir": parsed["owner_direct_ir"],
        "runner_ancestry": runner_path,
        "setup_ancestry": setup_path,
        "functions": parsed["functions"],
        "incoming_edges": parsed["incoming_edges"],
        "file": parsed["file"],
        "sha256": parsed["sha256"],
        "warnings": parsed["warnings"],
        "validation": {
            **parsed["validation"],
            "positive_caller_classified_from_raw_edge": True,
            "positive_runner_or_setup_ancestry": True,
            "expected_xls_wrapper_or_direct_runner": (
                job["shape"] is not None or caller_matches(caller, job["runner"])
                or caller_matches(caller, XLS_OWNER_WRAPPER)
            ),
            "single_constructor_call": True,
        },
    }


def numbered_paths(stage_dir: Path, name: str) -> list[tuple[int, Path]]:
    prefix = f"{name}.callgrind."
    found: list[tuple[int, Path]] = []
    for path in stage_dir.glob(prefix + "*"):
        suffix = path.name[len(prefix):]
        if suffix.isdigit() and path.is_file():
            found.append((int(suffix), path))
    found.sort()
    require(found, f"{stage_dir.name}/{name}: no numbered Callgrind dumps")
    require([number for number, _ in found] == list(range(1, len(found) + 1)),
            f"{stage_dir.name}/{name}: numbered dumps are not contiguous")
    return found


def final_dump(path: Path, expected_part: int, owner: str) -> dict[str, Any]:
    text = read_text(path)
    label = relative(path)
    require(part_number(text, label) == expected_part,
            f"{label}: final dump part differs from numbered dump count")
    require(trigger(text, label) == "Program termination",
            f"{label}: final dump is not a Program termination dump")
    require("events: Ir" in text.splitlines(), f"{label}: final event set is not Ir")
    values = []
    for line in text.splitlines():
        match = SUMMARY_RE.match(line.strip())
        if match:
            for token in match.group(1).split():
                try:
                    values.append(int(token.replace(",", "")))
                except ValueError:
                    break
            break
    require(values and values[0] == 0,
            f"{label}: final process dump Ir is {values[0] if values else None}, expected zero")
    return {
        "file": label,
        "sha256": sha256(path),
        "part": expected_part,
        "trigger": "Program termination",
        "summary_ir": 0,
        "owner": owner,
        "validation": {"final_process_dump_zero_ir": True},
    }


def annotation_command(path: Path, inclusive: bool) -> list[str]:
    return [
        "callgrind_annotate", "--auto=no", "--threshold=100", "--show-percs=no",
        "--inclusive=" + ("yes" if inclusive else "no"), "--tree=both", str(path),
    ]


def run_annotation(path: Path, inclusive: bool) -> tuple[str, list[str]]:
    command = annotation_command(path, inclusive)
    environment = dict(os.environ)
    environment.update(PERL_ENV)
    try:
        process = subprocess.run(command, capture_output=True, text=True,
                                 check=True, env=environment)
    except (OSError, subprocess.CalledProcessError) as error:
        raise EvidenceError(f"callgrind_annotate failed for {relative(path)}: {error}") from error
    require(process.stderr == "",
            f"callgrind_annotate emitted stderr for {relative(path)}")
    return process.stdout, command


def display_name(text: str) -> str:
    return EDGES.display_name(text)


def parse_annotation(text: str, selected: str, label: str) -> dict[str, Any]:
    rows = []
    lines = text.splitlines()
    for index, line in enumerate(lines):
        match = EDGES.STAR_RE.match(line)
        if match and display_name(match.group(2)) == selected:
            rows.append((index, int(match.group(1).replace(",", ""))))
    require(len(rows) == 1,
            f"{label}: expected one annotation row for {selected}, got {rows}")
    row_index, selected_ir = rows[0]
    direct: list[dict[str, Any]] = []
    for line in lines[row_index + 1:]:
        match = EDGES.EDGE_RE.match(line)
        if not match:
            break
        rendered_name = match.group(2)
        calls = re.search(r"\(([\d,]+)x\)", rendered_name)
        direct.append({
            "name": display_name(rendered_name),
            "inclusive_ir": int(match.group(1).replace(",", "")),
            "calls": int(calls.group(1).replace(",", "")) if calls else None,
        })
    return {"selected_ir": selected_ir, "direct": direct}


def annotate_timed_dump(stage_dir: Path, name: str, dump: Path,
                         expected_total: int, owner: str) -> dict[str, Any]:
    inclusive_text, inclusive_command = run_annotation(dump, True)
    self_text, self_command = run_annotation(dump, False)
    inclusive_path = stage_dir / f"{name}.part-{dump.name.rsplit('.', 1)[-1]}.inclusive.txt"
    self_path = stage_dir / f"{name}.part-{dump.name.rsplit('.', 1)[-1]}.self.txt"
    inclusive_path.write_text(inclusive_text, encoding="utf-8")
    self_path.write_text(self_text, encoding="utf-8")
    inclusive = parse_annotation(inclusive_text, owner, relative(inclusive_path))
    exclusive = parse_annotation(self_text, owner, relative(self_path))
    require(inclusive["selected_ir"] == expected_total,
            f"{relative(inclusive_path)}: owner Ir differs from raw incoming edge")
    require(inclusive["direct"] == exclusive["direct"],
            f"{relative(self_path)}: visible direct children differ")
    require(exclusive["selected_ir"] + sum(
        edge["inclusive_ir"] for edge in exclusive["direct"]
    ) == inclusive["selected_ir"],
            f"{relative(self_path)}: self plus visible direct Ir differs")
    return {
        "dump": relative(dump),
        "environment": dict(PERL_ENV),
        "command": {
            "inclusive": [*inclusive_command[:-1], relative(dump)],
            "self": [*self_command[:-1], relative(dump)],
        },
        "files": {
            "inclusive": relative(inclusive_path),
            "inclusive_sha256": sha256(inclusive_path),
            "self": relative(self_path),
            "self_sha256": sha256(self_path),
        },
        "owner": {
            "inclusive_ir": inclusive["selected_ir"],
            "self_ir": exclusive["selected_ir"],
            "direct_callees": inclusive["direct"],
        },
        "validation": {
            "inclusive_owner_matches_raw": True,
            "self_plus_visible_direct_matches_inclusive": True,
            "inclusive_and_self_direct_children_match": True,
            "deterministic_perl_environment": True,
        },
    }


def attribution_for(parsed: dict[str, Any], target: str, label: str) -> dict[str, Any]:
    functions = parsed["functions"]
    ids = [function_id for function_id, function in functions.items()
           if target in function["name"]]
    require(ids, f"{label}: attribution target {target!r} is absent")
    selected = set(ids)
    incoming = []
    for parent_id, function in functions.items():
        for edge in function["edges"]:
            if edge["callee_id"] in selected and edge["calls"] is not None \
                    and edge["calls"] > 0 and edge["inclusive_ir"] > 0:
                incoming.append({
                    "caller_id": parent_id,
                    "caller": function["name"],
                    "callee_id": edge["callee_id"],
                    "calls": edge["calls"],
                    "inclusive_ir": edge["inclusive_ir"],
                    "line": edge["line"],
                })
    incoming.sort(key=lambda edge: (
        edge["caller"], edge["callee_id"], edge["line"] or 0,
    ))
    require(incoming, f"{label}: attribution target {target!r} has no positive incoming edge")
    self_ir = sum(functions[function_id]["self_ir"] for function_id in ids)
    direct_ir = sum(
        edge["inclusive_ir"]
        for function_id in ids
        for edge in functions[function_id]["edges"]
    )
    return {
        "target": target,
        "function_ids": ids,
        "self_ir": self_ir,
        "direct_ir": direct_ir,
        "inclusive_ir": sum(edge["inclusive_ir"] for edge in incoming),
        "calls": sum(edge["calls"] for edge in incoming),
        "incoming_edge_count": len(incoming),
        "incoming_edges": incoming,
        "validation": {
            "target_present": True,
            "positive_incoming_edges_retained": True,
        },
    }


def profile_command(plan: dict[str, Any], stage_dir: Path, job: dict[str, Any],
                    receipt: dict[str, Any], metadata: dict[str, Any]) -> None:
    label = f"baseline/{job['name']}.receipt.json"
    require(receipt.get("exit_code") == 0, f"{label}: profile child failed")
    require(receipt.get("binary_sha256") == metadata["binary_sha256"],
            f"{label}: profile binary binding differs")
    require(receipt.get("source_manifest_sha256") == metadata["source_manifest_sha256"],
            f"{label}: profile source binding differs")
    require(receipt.get("plan_sha256") == sha256(PLAN_PATH),
            f"{label}: profile plan binding differs")
    require(receipt.get("script_sha256") == sha256(RUN_PATH),
            f"{label}: profile run binding differs")
    command = receipt.get("command")
    require(isinstance(command, list) and all(isinstance(item, str) for item in command),
            f"{label}: command is not an argv list")
    require(command[:3] == ["taskset", "-c", str(plan["cpu"])],
            f"{label}: profile is not pinned to the plan CPU")
    require(command[3:6] == ["valgrind", "--tool=callgrind", "--collect-atstart=no"],
            f"{label}: profile collection mode differs")
    require(f"--toggle-collect={job['owner']}" in command and
            f"--zero-before={job['owner']}" in command and
            f"--dump-after={job['owner']}" in command,
            f"{label}: profile does not bind all exact owner controls")
    require(option(command, "--callgrind-out-file") == str(stage_dir / f"{job['name']}.callgrind"),
            f"{label}: Callgrind output path differs")
    require(str(metadata["binary_path"]) in command,
            f"{label}: profile binary path differs")
    require(option(command, "--case") == ",".join(job["selection"]["cases"]),
            f"{label}: selected case differs")
    require(option(command, "--warmup") == "0", f"{label}: warmup differs")
    require(option(command, "--samples") == "5", f"{label}: sample count differs")
    require(option(command, "--json") == str(stage_dir / f"{job['name']}.json"),
            f"{label}: report path differs")
    require(option(command, "--corpus-manifest") == str(stage_dir / f"{job['name']}.catalog.json"),
            f"{label}: catalog path differs")
    if job["shape"] is None:
        require("--shape" not in command and "--payload" not in command,
                f"{label}: XLS profile unexpectedly carries CFB selectors")
    else:
        require(option(command, "--shape") == job["shape"],
                f"{label}: CFB shape differs")
        require(option(command, "--payload") == "incompressible",
                f"{label}: CFB payload differs")


def normalize(value: Any, label: str) -> Any:
    """Collapse constant report vectors while rejecting silent variation."""

    if isinstance(value, list):
        require(value, f"{label}: empty identity vector")
        normalized = [normalize(item, f"{label}[{index}]")
                      for index, item in enumerate(value)]
        require(all(item == normalized[0] for item in normalized),
                f"{label}: identity vector varies across samples")
        return normalized[0]
    if isinstance(value, dict):
        return {key: normalize(item, f"{label}.{key}")
                for key, item in sorted(value.items())}
    return value


def identity(row: dict[str, Any], label: str) -> dict[str, Any]:
    require(isinstance(row, dict), f"{label}: result row is not an object")
    return {
        "case": row.get("case"),
        "corpus": row.get("corpus"),
        "sink": row.get("sink"),
        "source": normalize(row.get("source"), label + ".source")
                      if row.get("source") is not None else None,
        "output_sha256": row.get("output_sha256"),
    }


def validate_numeric_receipt(numeric: Any, name: str) -> dict[str, Any]:
    """Run the campaign verifier's receipt/hash checks for one child."""

    try:
        receipt = numeric.receipt(name, "normal")
    except (AssertionError, KeyError, OSError, TypeError, ValueError) as error:
        raise EvidenceError(f"baseline/{name}.receipt.json: receipt validation failed: {error}") from error
    require(isinstance(receipt, dict), f"baseline/{name}.receipt.json: receipt is not an object")
    return receipt


def validate_profile_report(numeric: Any, stage_dir: Path, job: dict[str, Any],
                            metadata: dict[str, Any]) -> tuple[dict[str, Any], dict[str, Any]]:
    path = stage_dir / f"{job['name']}.json"
    require(path.is_file(), f"{relative(path)}: profile report is missing")
    try:
        rows = numeric.validate_report(path, job, "normal")
    except (AssertionError, KeyError, OSError, TypeError, ValueError) as error:
        raise EvidenceError(f"{relative(path)}: numerical report validation failed: {error}") from error
    report = read_json(path)
    require(len(rows) == 1, f"{relative(path)}: expected one profile result")
    row = rows[0]
    require(row.get("case") == job["selection"]["cases"][0],
            f"{relative(path)}: profile case differs")
    if job["shape"] is not None:
        require(row.get("corpus", {}).get("shape") == job["shape"],
                f"{relative(path)}: profile corpus shape differs")
    require(report.get("binary_identity", {}).get("binary_sha256") == metadata["binary_sha256"],
            f"{relative(path)}: profile binary identity differs")
    return report, row


def validate_native_counterpart(numeric: Any, stage_dir: Path, plan: dict[str, Any],
                                job: dict[str, Any], metadata: dict[str, Any],
                                cache: dict[str, tuple[list[dict[str, Any]], dict[str, Any]]]) -> tuple[dict[str, Any], dict[str, Any]]:
    group = "xls" if job["shape"] is None else "cfb"
    native_name = f"native-r{job['repeat']}-{group}"
    if native_name not in cache:
        native_job = next(item for item in native_jobs(plan) if item["name"] == native_name)
        report_path = stage_dir / f"{native_name}.json"
        receipt_path = stage_dir / f"{native_name}.receipt.json"
        require(receipt_path.is_file(), f"{relative(receipt_path)}: native receipt is missing")
        require(report_path.is_file(), f"{relative(report_path)}: native report is missing")
        try:
            receipt = validate_numeric_receipt(numeric, native_name)
            rows = numeric.validate_report(report_path, native_job, "normal")
        except (AssertionError, KeyError, OSError, TypeError, ValueError) as error:
            raise EvidenceError(f"{native_name}: native report validation failed: {error}") from error
        require(receipt.get("binary_sha256") == metadata["binary_sha256"],
                f"{native_name}: native binary binding differs")
        rss_path = stage_dir / f"{native_name}.rss.json"
        require(rss_path.is_file(), f"{relative(rss_path)}: native RSS evidence is missing")
        rss = read_json(rss_path)
        require(rss.get("max_rss_kib", 0) > 0,
                f"{relative(rss_path)}: native RSS is not positive")
        cache[native_name] = (rows, {
            "report": relative(report_path),
            "report_sha256": sha256(report_path),
            "receipt": relative(receipt_path),
            "receipt_sha256": sha256(receipt_path),
            "rss": relative(rss_path),
            "max_rss_kib": rss["max_rss_kib"],
        })
    rows, native_meta = cache[native_name]
    matches = [row for row in rows
               if row.get("case") == job["selection"]["cases"][0]
               and (job["shape"] is None or row.get("corpus", {}).get("shape") == job["shape"])]
    require(len(matches) == 1, f"{native_name}: expected one native counterpart row")
    return native_meta, matches[0]


def analyze_profile(plan: dict[str, Any], metadata: dict[str, Any], numeric: Any,
                    job: dict[str, Any], native_cache: dict[str, Any]) -> dict[str, Any]:
    stage_dir = HERE / "baseline"
    name = job["name"]
    receipt_path = stage_dir / f"{name}.receipt.json"
    require(receipt_path.is_file(), f"{relative(receipt_path)}: profile receipt is missing")
    receipt = read_json(receipt_path)
    validate_numeric_receipt(numeric, name)
    profile_command(plan, stage_dir, job, receipt, metadata)
    numbers = numbered_paths(stage_dir, name)
    expected_artifacts = {
        f"{name}.json", f"{name}.catalog.json", f"{name}.stdout", f"{name}.stderr",
        f"{name}.callgrind",
        *(f"{name}.callgrind.{number}" for number, _ in numbers),
    }
    validate_receipt_artifacts(stage_dir, receipt, expected_artifacts,
                               f"baseline/{name}")
    dumps = [validate_numbered_dump(path, number, job, plan)
             for number, path in numbers]
    expected_setup = 0 if job["shape"] is None else 1
    setup = [dump for dump in dumps if dump["role"] == "setup"]
    timed = [dump for dump in dumps if dump["role"] == "timed"]
    require(len(setup) == expected_setup,
            f"{name}: raw caller proof found {len(setup)} setup dumps, expected {expected_setup}")
    require(len(timed) == plan["profile"]["samples"],
            f"{name}: raw caller proof found {len(timed)} timed dumps, expected five")
    final_path = stage_dir / f"{name}.callgrind"
    require(final_path.is_file(), f"{relative(final_path)}: final dump is missing")
    final = final_dump(final_path, len(dumps) + 1, job["owner"])

    profile_report, profile_row = validate_profile_report(numeric, stage_dir, job, metadata)
    native_meta, native_row = validate_native_counterpart(
        numeric, stage_dir, plan, job, metadata, native_cache
    )
    profile_identity = identity(profile_row, f"{name}.profile")
    native_identity = identity(native_row, f"{name}.native")
    require(profile_identity == native_identity,
            f"{name}: profile/native corpus, output, sink, or source identity differs")

    annotations = []
    attribution = []
    for dump in timed:
        dump_path = HERE / dump["file"]
        annotations.append(annotate_timed_dump(
            stage_dir, name, dump_path, dump["summary_ir"], job["owner"]
        ))
        parsed = parse_raw_profile(dump_path, job["owner"])
        functions = {
            label: attribution_for(parsed, target, f"{name}.{dump['number']}")
            for label, target in ATTRIBUTION_TARGETS.items()
        }
        attribution.append({
            "dump": dump["file"],
            "constructor": {
                "owner": job["owner"],
                "caller": dump["owner_caller"],
                "calls": dump["owner_calls"],
                "inclusive_ir": dump["summary_ir"],
                "self_ir": dump["owner_self_ir"],
                "direct_ir": dump["owner_direct_ir"],
            },
            "functions": functions,
            "validation": {
                "one_positive_timed_runner_edge": True,
                "one_selected_constructor_call": dump["owner_calls"] == 1,
                "all_requested_chain_targets_present": True,
            },
        })

    profile_stderr = stage_dir / f"{name}.stderr"
    warnings = {
        "stderr": {
            "file": relative(profile_stderr),
            "sha256": sha256(profile_stderr),
            "lines": read_text(profile_stderr).splitlines(),
        },
        "raw_dump_warning_lines": [line for dump in dumps
                                    for line in dump["warnings"]],
        "collection_off_call_metadata_caveat": (
            "Callgrind collection is disabled outside the exact owner toggle. "
            "Child-call metadata in the collected dump can retain context from "
            "collection-off caller paths; only the positive incoming owner edge "
            "and its matching dump summary establish timed scope."
        ),
    }
    return {
        "name": name,
        "repeat": job["repeat"],
        "group": job["group"],
        "shape": job["shape"],
        "owner": job["owner"],
        "receipt": relative(receipt_path),
        "receipt_sha256": sha256(receipt_path),
        "profile_result": relative(stage_dir / f"{name}.json"),
        "profile_result_sha256": sha256(stage_dir / f"{name}.json"),
        "native": native_meta,
        "profile_native_identity": profile_identity,
        "raw_dumps": dumps,
        "setup_dumps": setup,
        "timed_dumps": timed,
        "final_process_dump": final,
        "annotations": annotations,
        "constructor_attribution": attribution,
        "warnings": warnings,
        "validation": {
            "receipt_and_command": True,
            "profile_native_identity_parity": True,
            "raw_positive_caller_scope_proof": True,
            "setup_excluded_from_timed_attribution": len(setup) == expected_setup,
            "five_single_constructor_timed_dumps": len(timed) == 5 and
                                                     all(dump["owner_calls"] == 1 for dump in timed),
            "final_process_dump_zero_ir": True,
            "deterministic_annotations": True,
            "requested_cfb_chain_attribution_present": True,
        },
        "limitations": {
            "no_speedup_comparison": True,
            "callgrind_ir_is_mechanism_diagnostic": True,
            "no_native_latency_or_hardware_claim": True,
            "no_allocation_count_claim": True,
        },
    }


def analyze(stage_selection: str = "baseline") -> dict[str, Any]:
    plan = plan_data()
    require(stage_selection == "baseline",
            "0523 profile analysis is baseline attribution only; no candidate comparison exists")
    metadata = validate_build(plan)
    numeric = load_numeric_helper()
    native_cache: dict[str, Any] = {}
    profiles = [analyze_profile(plan, metadata, numeric, job, native_cache)
                for job in profile_jobs(plan)]
    require(len(profiles) == 8, "profile matrix is incomplete")
    identities: dict[tuple[str, str | None], dict[str, Any]] = {}
    for profile in profiles:
        key = profile["group"], profile["shape"]
        current = profile["profile_native_identity"]
        if key in identities:
            require(identities[key] == current,
                    f"profile identity changed across repeats for {key}")
        else:
            identities[key] = current
    return {
        "schema": "cfb_ole2_constructor_callgrind_profile_analysis_v1",
        "status": "pass",
        "scope": plan["scope"],
        "performance_claim": "none",
        "selected_owners": {
            "xls": plan["profile"]["xls_owner"],
            "cfb": plan["profile"]["cfb_owner"],
        },
        "plan": relative(PLAN_PATH),
        "plan_sha256": sha256(PLAN_PATH),
        "stage_selection": ["baseline"],
        "metadata": metadata,
        "profile_count": len(profiles),
        "timed_constructor_dump_count": sum(len(item["timed_dumps"]) for item in profiles),
        "setup_dump_count": sum(len(item["setup_dumps"]) for item in profiles),
        "profiles": profiles,
        "native_counterpart_reports": sorted(
            (meta for _rows, meta in native_cache.values()),
            key=lambda item: item["report"],
        ),
        "helpers": {
            "change-0519/analyze_profiles.py": sha256(RETAINED_DIR / "analyze_profiles.py"),
            "change-0519/compare_profile_lanes.py": sha256(
                RETAINED_DIR / "compare_profile_lanes.py"
            ),
            "analyze.py": sha256(ROOT_ANALYZER),
        },
        "comparison": None,
        "validation": {
            "source_binary_receipt_hashes_bound": True,
            "all_profile_receipts_and_artifacts_valid": True,
            "all_profile_native_output_identity_valid": True,
            "raw_positive_incoming_caller_edges_valid": True,
            "setup_and_timed_dumps_classified_from_edges": True,
            "all_selected_timed_dumps_single_constructor": True,
            "all_final_process_dumps_zero_ir": True,
            "all_requested_chain_functions_summarized": True,
            "no_speedup_comparison_performed": True,
        },
        "limitations": (
            "This is baseline CFB/OLE2 constructor attribution only. Callgrind Ir, "
            "inclusive rows, and raw child-call metadata are mechanism diagnostics; "
            "they do not establish native latency, hardware instructions/cycles, "
            "allocation counts, RSS, physical I/O, cold-cache behavior, scaling, "
            "or native Office-producer behavior."
        ),
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("output", nargs="?", type=Path,
                        help="JSON destination; defaults to profile-analysis.json")
    parser.add_argument("--output", dest="output_option", type=Path,
                        help="JSON destination (alternative to positional path)")
    parser.add_argument("--stage", choices=("baseline",), default="baseline",
                        help="validate the baseline attribution stage")
    args = parser.parse_args()
    if args.output is not None and args.output_option is not None:
        parser.error("provide the output path either positionally or with --output")
    output = args.output_option or args.output or (HERE / "profile-analysis.json")
    try:
        document = analyze(args.stage)
        output.parent.mkdir(parents=True, exist_ok=True)
        output.write_text(json.dumps(document, indent=2) + "\n", encoding="utf-8")
    except (EvidenceError, OSError, json.JSONDecodeError) as error:
        print(f"analyze_profiles.py: error: {error}", file=sys.stderr)
        return 2
    print(f"CFB/OLE2 baseline constructor profiles verified: {output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
