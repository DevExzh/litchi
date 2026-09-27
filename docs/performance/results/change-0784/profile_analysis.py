#!/usr/bin/env python3
"""Offline validation and attribution of the 0784 Callgrind captures.

The profile driver retains one positive numbered Callgrind dump and one zero
program-termination dump for each shape/pass.  This module only reads those
artifacts.  It resolves Callgrind's compressed function names before parsing
the records, accepts absolute, relative and wildcard source positions, and
reconstructs the selected wrapper from its raw incoming edge.  The selected
function's own Ir and its immediate children's inclusive Ir form the only
disjoint partition.  Descendant inclusive rows are retained as a path
diagnostic and are never added to that partition.

Callgrind Ir is a guest-instruction diagnostic.  This report does not claim
native latency, a production speedup, or a causal explanation of the 0780
measurement.
"""

from __future__ import annotations

import argparse
import collections
import hashlib
import json
import re
import sys
from pathlib import Path
from typing import Any, NoReturn


HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[3]
OLD_PACKET = HERE.parent / "change-0780"
OWNER = "pptx_capture_probe::capture_region_0784"
CHILD_EXPECTED = (
    "litchi_pptx::package::model::Package::opened_presentation_with_limits"
)
NESTED_EXPECTED = "litchi_pptx::opened::model::capture_internal"
MARKER = "litchi-perf-0780-static-mce-capabilities"
SHAPES = ("tiny", "medium", "large")
DIMENSIONS = {"tiny": (3, 4), "medium": (12, 8), "large": (100, 100)}
PROFILE_ORDER = (("tiny", "medium", "large"), ("large", "medium", "tiny"))
ANALYSIS_JSON = HERE / "profile-analysis.json"
ANALYSIS_MD = HERE / "profile-analysis.md"

FUNCTION_RE = re.compile(r"^(fn|cfn)=\((\d+)\)(?:\s+(.*))?$")
CALLS_RE = re.compile(r"^calls=([+\-]?[0-9][0-9,]*)")
SUMMARY_RE = re.compile(r"^summary:\s*(.*)$")
TOTALS_RE = re.compile(r"^totals:\s*(.*)$")
PART_RE = re.compile(r"^part:\s*(\d+)$")
TRIGGER_RE = re.compile(r"^desc:\s+Trigger:\s+(.*)$")
CMD_RE = re.compile(r"^cmd:\s?(.*)$")
EVENTS_RE = re.compile(r"^events:\s*(.*)$")


class EvidenceError(ValueError):
    """A missing, malformed, or contradictory retained artifact."""


def fail(message: str) -> NoReturn:
    raise EvidenceError(message)


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    try:
        with path.open("rb") as stream:
            for chunk in iter(lambda: stream.read(1024 * 1024), b""):
                digest.update(chunk)
    except OSError as error:
        fail(f"cannot hash {path}: {error}")
    return digest.hexdigest()


def read_json(path: Path) -> Any:
    require(path.is_file() and not path.is_symlink(), f"missing JSON: {path}")
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        fail(f"invalid JSON {path}: {error}")


def read_text(path: Path) -> str:
    try:
        return path.read_text(encoding="utf-8", errors="replace")
    except OSError as error:
        fail(f"cannot read {path}: {error}")


def relative(path: Path) -> str:
    try:
        return str(path.relative_to(HERE))
    except ValueError as error:
        fail(f"path escapes packet: {path}")
        raise AssertionError from error


def json_digest(value: Any) -> str:
    encoded = (json.dumps(value, sort_keys=True, separators=(",", ":")) + "\n").encode()
    return hashlib.sha256(encoded).hexdigest()


def packet_path(raw: str, label: str) -> Path:
    """Resolve a retained absolute/relative path after packet relocation."""

    require(isinstance(raw, str) and raw, f"{label}: path is missing")
    marker = "/change-0784/"
    if marker in raw:
        path = HERE / raw.split(marker, 1)[1]
    else:
        candidate = Path(raw)
        path = candidate if candidate.is_absolute() else HERE / candidate
    path = path.resolve()
    require(path == HERE or HERE in path.parents, f"{label}: path escapes packet: {raw}")
    return path


def external_path(raw: str, label: str) -> Path:
    """Resolve an owned executable path, which intentionally lives outside the packet."""

    require(isinstance(raw, str) and raw, f"{label}: path is missing")
    marker = "/change-0784/"
    if marker in raw:
        return (HERE / raw.split(marker, 1)[1]).resolve()
    path = Path(raw)
    require(path.is_absolute(), f"{label}: executable path must be absolute")
    return path.resolve()


def artifact(value: Any, label: str, allow_missing: bool = False) -> dict[str, Any]:
    require(isinstance(value, dict), f"{label}: artifact is not an object")
    path = packet_path(value.get("path"), label)
    expected_sha = value.get("sha256")
    expected_bytes = value.get("bytes")
    require(isinstance(expected_sha, str) and re.fullmatch(r"[0-9a-f]{64}", expected_sha),
            f"{label}: artifact SHA-256 is invalid")
    require(type(expected_bytes) is int and expected_bytes >= 0,
            f"{label}: artifact byte count is invalid")
    if not path.is_file() or path.is_symlink():
        if allow_missing:
            return {"path": relative(path), "sha256": expected_sha,
                    "bytes": expected_bytes, "custody": "cleanup-witness"}
        fail(f"{label}: artifact is missing: {path}")
    require(path.stat().st_size == expected_bytes,
            f"{label}: artifact byte count changed")
    require(sha256(path) == expected_sha, f"{label}: artifact SHA-256 changed")
    return {"path": relative(path), "sha256": expected_sha,
            "bytes": expected_bytes, "custody": "live"}


def validate_plan() -> dict[str, Any]:
    plan = read_json(HERE / "plan.json")
    require(isinstance(plan, dict), "plan is not an object")
    require(plan.get("schema") == "litchi.performance.0784.v1", "plan schema changed")
    require(plan.get("cpu") == 12, "profile CPU changed")
    require(plan.get("owner") == OWNER, "profile owner changed")
    require(plan.get("scope") and "wrapper" in plan["scope"].lower(),
            "profile scope is missing")
    profile = plan.get("profile")
    require(isinstance(profile, dict), "profile plan is missing")
    require(profile.get("collect_at_start") is False, "Callgrind collection policy changed")
    require(profile.get("events") == ["Ir"], "profile event set changed")
    require(profile.get("repeats") == 2 and profile.get("samples") == 1
            and profile.get("warmup") == 0, "profile repeat/sample policy changed")
    require(profile.get("orders") == [list(order) for order in PROFILE_ORDER],
            "profile shape order changed")
    require(profile.get("expected_numbered_parts") == 1,
            "numbered Callgrind part policy changed")
    return plan


def cleanup_binary_witness(binary: dict[str, Any], label: str) -> dict[str, Any]:
    """Validate an executable while it is live or after exact cleanup."""

    path = external_path(binary.get("path"), label)
    expected_sha = binary.get("sha256")
    expected_bytes = binary.get("bytes")
    require(isinstance(expected_sha, str) and len(expected_sha) == 64,
            f"{label}: binary hash is missing")
    require(type(expected_bytes) is int and expected_bytes > 0,
            f"{label}: binary byte count is invalid")
    if path.is_file() and not path.is_symlink():
        require(path.stat().st_size == expected_bytes, f"{label}: binary size changed")
        require(sha256(path) == expected_sha, f"{label}: binary hash changed")
        # Keep custody mode out of the generated report.  The live executable
        # and its later cleanup witness must replay to the same identity.
        return {"path": str(path), "sha256": expected_sha,
                "bytes": expected_bytes, "verified": True}

    cleanup_path = HERE / "cleanup.json"
    require(cleanup_path.is_file(), f"{label}: missing binary lacks cleanup witness")
    cleanup = read_json(cleanup_path)
    require(cleanup.get("verified_before_removal") is True,
            f"{label}: cleanup witness is not verified")
    witnesses = cleanup.get("binaries")
    require(isinstance(witnesses, dict), f"{label}: cleanup binary map is missing")
    key = Path(str(binary["path"])).name
    witness = witnesses.get(key) or witnesses.get(label.rsplit(".", 1)[-1])
    require(isinstance(witness, dict), f"{label}: cleanup witness does not bind binary")
    require(witness.get("sha256") == expected_sha
            and witness.get("bytes") == expected_bytes,
            f"{label}: cleanup binary identity differs")
    return {"path": str(path), "sha256": expected_sha,
            "bytes": expected_bytes, "verified": True}


def validate_build(plan: dict[str, Any]) -> dict[str, Any]:
    path = HERE / "build" / "build.json"
    build = read_json(path)
    require(isinstance(build, dict), "build manifest is not an object")
    binaries = build.get("binaries")
    require(isinstance(binaries, dict) and set(binaries) == {"control", "profile"},
            "build binary matrix changed")
    for row in build.get("commands", []):
        require(isinstance(row, dict) and row.get("exit_code") == 0,
                "a recorded build command failed")
        log = row.get("log")
        if isinstance(log, dict):
            artifact(log, "build log")
    source_ref = build.get("source")
    require(isinstance(source_ref, dict), "build source identity is missing")
    source_path = packet_path(source_ref.get("path"), "build source")
    require(source_path == HERE / "build" / "source.json", "build source path changed")
    artifact(source_ref, "build source")
    probe = build.get("probe")
    require(isinstance(probe, dict) and probe, "probe manifest is missing")
    for name, digest in probe.items():
        path_for_probe = HERE / name
        require(path_for_probe.is_file() and not path_for_probe.is_symlink(),
                f"probe source missing: {name}")
        require(sha256(path_for_probe) == digest, f"probe source hash changed: {name}")
    custody = {
        "build": relative(path),
        "build_sha256": sha256(path),
        "source": artifact(source_ref, "build source"),
        "binaries": {},
    }
    for name in ("control", "profile"):
        row = binaries[name]
        custody["binaries"][name] = cleanup_binary_witness(row, f"build {name} binary")
    return {"manifest": build, "manifest_sha256": sha256(path), "custody": custody}


def receipt_artifacts(receipt: dict[str, Any], stem: str) -> dict[str, Path]:
    values = receipt.get("artifacts")
    require(isinstance(values, dict), f"{stem}: receipt artifacts are missing")
    expected = {f"{stem}.json", f"{stem}.log", f"{stem}.callgrind",
                f"{stem}.callgrind.1"}
    require(set(values) == expected,
            f"{stem}: artifact set differs from {sorted(expected)}")
    result: dict[str, Path] = {}
    for name in sorted(expected):
        descriptor = values[name]
        require(Path(name).name == name, f"{stem}: artifact key escapes packet")
        result[name] = packet_path(descriptor.get("path"), f"{stem} {name}")
        artifact(descriptor, f"{stem} {name}")
    return result


def command_path(raw: str, label: str) -> Path:
    return packet_path(raw, label)


def validate_profile_receipt(
    row: dict[str, Any], plan: dict[str, Any], build: dict[str, Any],
    block: int, shape: str,
) -> dict[str, Any]:
    stem = f"{block}-{shape}"
    label = f"profiles/{stem}"
    require(row.get("block") == block and row.get("shape") == shape,
            f"{label}: receipt matrix position changed")
    require(row.get("exit_code") == 0, f"{label}: profiler child failed")
    require(row.get("binary") == build["manifest"]["binaries"]["profile"],
            f"{label}: binary identity differs from profile build")
    require(row.get("driver_sha256") == sha256(HERE / "profile.py"),
            f"{label}: profile driver hash changed")
    command = row.get("command")
    require(isinstance(command, list), f"{label}: command is not a list")
    fixed = [
        "taskset", "-c", str(plan["cpu"]), "valgrind", "--tool=callgrind",
        "--collect-atstart=no", "--toggle-collect=" + OWNER,
        "--zero-before=" + OWNER, "--dump-after=" + OWNER,
    ]
    require(command[: len(fixed)] == fixed, f"{label}: Callgrind command prefix changed")
    require(len(command) == 21, f"{label}: Callgrind command length changed")
    raw_path = command[9]
    require(isinstance(raw_path, str) and raw_path.startswith("--callgrind-out-file="),
            f"{label}: Callgrind output option is missing")
    raw_output = raw_path.split("=", 1)[1]
    require(command_path(raw_output, f"{label} Callgrind output")
            == HERE / "profiles" / f"{stem}.callgrind",
            f"{label}: Callgrind output path changed")
    binary_path = external_path(command[10], f"{label} binary")
    expected_binary_path = external_path(
        build["manifest"]["binaries"]["profile"]["path"], f"{label} build binary"
    )
    require(binary_path == expected_binary_path, f"{label}: binary path changed")
    expected_args = ["--mode", "capture", "--shape", shape, "--samples", "1",
                     "--warmup", "0", "--output"]
    # The explicit positions keep path relocation separate from semantic args.
    require(command[11:20] == expected_args,
            f"{label}: capture arguments changed")
    output_path = command_path(command[20], f"{label} JSON output")
    require(output_path == HERE / "profiles" / f"{stem}.json",
            f"{label}: JSON output path changed")
    paths = receipt_artifacts(row, stem)
    return {"stem": stem, "paths": paths, "command": list(command),
            "binary": build["manifest"]["binaries"]["profile"]}


def header_values(lines: list[str]) -> dict[str, Any]:
    events: list[str] | None = None
    summary: int | None = None
    totals: int | None = None
    part: int | None = None
    trigger: str | None = None
    command: str | None = None
    for line in lines:
        stripped = line.strip()
        if match := EVENTS_RE.match(stripped):
            require(events is None, "Callgrind has duplicate events headers")
            events = match.group(1).split()
        elif match := SUMMARY_RE.match(stripped):
            values = parse_ints(match.group(1))
            require(values, "Callgrind summary is not numeric")
            require(summary is None, "Callgrind has duplicate summaries")
            summary = values[0]
        elif match := TOTALS_RE.match(stripped):
            values = parse_ints(match.group(1))
            require(values, "Callgrind totals is not numeric")
            require(totals is None, "Callgrind has duplicate totals")
            totals = values[0]
        elif match := PART_RE.match(stripped):
            require(part is None, "Callgrind has duplicate parts")
            part = int(match.group(1))
        elif match := TRIGGER_RE.match(stripped):
            require(trigger is None, "Callgrind has duplicate triggers")
            trigger = match.group(1)
        elif match := CMD_RE.match(stripped):
            if command is None:
                command = match.group(1)
    require(events is not None and events == ["Ir"], "Callgrind event set is not exactly Ir")
    require(summary is not None and summary >= 0, "Callgrind summary is missing")
    require(part is not None, "Callgrind part is missing")
    require(trigger is not None, "Callgrind trigger is missing")
    return {"events": events, "summary_ir": summary, "totals_ir": totals,
            "part": part, "trigger": trigger, "command": command}


def parse_ints(text: str) -> list[int]:
    values: list[int] = []
    for token in text.split():
        try:
            values.append(int(token.replace(",", "")))
        except ValueError:
            break
    return values


def cost_line(line: str, event_index: int) -> tuple[int, str] | None:
    fields = line.split()
    if len(fields) <= event_index + 1:
        return None
    position = fields[0]
    if position != "*" and not re.fullmatch(r"[+-]?\d+", position):
        return None
    try:
        return int(fields[event_index + 1].replace(",", "")), position
    except ValueError:
        return None


def position_kind(position: str) -> str:
    if position == "*":
        return "wildcard"
    if position.startswith(("+", "-")):
        return "relative"
    return "absolute"


def add_name(names: dict[int, str], function_id: int, name: str) -> None:
    if not name:
        return
    prior = names.get(function_id)
    require(prior is None or prior == name,
            f"function id {function_id} has conflicting compressed names")
    names[function_id] = name


def parse_raw(path: Path, allow_empty: bool = False) -> dict[str, Any]:
    label = relative(path)
    lines = read_text(path).splitlines()
    header = header_values(lines)
    event_index = header["events"].index("Ir")
    names: dict[int, str] = {}
    declarations = 0
    for line in lines:
        match = FUNCTION_RE.match(line.strip())
        if match and match.group(3):
            declarations += 1
            add_name(names, int(match.group(2)), match.group(3))

    functions: dict[int, dict[str, Any]] = {}
    current_id: int | None = None
    pending: dict[str, Any] | None = None
    costs = 0
    edges = 0
    positions = collections.Counter[str]()
    for line_number, line in enumerate(lines, 1):
        stripped = line.strip()
        function_match = FUNCTION_RE.match(stripped)
        if function_match:
            kind, text_id, name = function_match.groups()
            function_id = int(text_id)
            if name:
                add_name(names, function_id, name)
            if kind == "fn":
                current_id = function_id
                pending = None
                function = functions.setdefault(
                    function_id,
                    {"id": function_id, "name": "", "records": 0,
                     "self_ir": 0, "edges": []},
                )
                function["records"] += 1
            else:
                require(current_id is not None,
                        f"{label}:{line_number}: cfn appears outside fn")
                pending = {"callee_id": function_id, "calls": None,
                           "line": line_number}
            continue
        calls_match = CALLS_RE.match(stripped)
        if calls_match:
            require(current_id is not None and pending is not None,
                    f"{label}:{line_number}: calls appears outside cfn")
            require(pending["calls"] is None,
                    f"{label}:{line_number}: duplicate calls record")
            pending["calls"] = int(calls_match.group(1).replace(",", ""))
            continue
        if current_id is None:
            continue
        parsed = cost_line(stripped, event_index)
        if parsed is None:
            continue
        value, position = parsed
        positions[position_kind(position)] += 1
        costs += 1
        function = functions[current_id]
        if pending is not None and pending["calls"] is not None:
            edge = {
                "caller_id": current_id,
                "callee_id": pending["callee_id"],
                "calls": pending["calls"],
                "inclusive_ir": value,
                "line": pending["line"],
                "position": position,
            }
            function["edges"].append(edge)
            edges += 1
            pending = None
        else:
            function["self_ir"] += value

    for function_id, function in functions.items():
        function["name"] = names.get(function_id, "")
    if allow_empty and header["summary_ir"] == 0 and not names and costs == 0:
        functions = {}
    else:
        require(names, f"{label}: no compressed function names were resolved")
        require(costs > 0, f"{label}: no Ir cost records were parsed")
    return {
        "path": label,
        "sha256": sha256(path),
        "bytes": path.stat().st_size,
        "header": header,
        "names": names,
        "functions": functions,
        "statistics": {
            "function_records": sum(f["records"] for f in functions.values()),
            "unique_functions": len(functions),
            "compressed_name_declarations": declarations,
            "cost_records": costs,
            "call_edge_records": edges,
            "position_kinds": dict(sorted(positions.items())),
        },
    }


def display_name(name: str) -> str:
    return name or "<unnamed>"


def aggregate_edges(function: dict[str, Any], names: dict[int, str]) -> list[dict[str, Any]]:
    grouped: dict[int, dict[str, Any]] = {}
    for edge in function["edges"]:
        if edge["calls"] <= 0 or edge["inclusive_ir"] <= 0:
            continue
        item = grouped.setdefault(edge["callee_id"], {
            "callee_id": edge["callee_id"],
            "callee": names.get(edge["callee_id"], ""),
            "calls": 0,
            "edge_count": 0,
            "inclusive_ir": 0,
        })
        item["calls"] += edge["calls"]
        item["edge_count"] += 1
        item["inclusive_ir"] += edge["inclusive_ir"]
    return sorted(grouped.values(), key=lambda item: (-item["inclusive_ir"], item["callee_id"]))


def function_view(function: dict[str, Any], names: dict[int, str]) -> dict[str, Any]:
    direct = aggregate_edges(function, names)
    return {
        "id": function["id"],
        "name": display_name(function["name"]),
        "records": function["records"],
        "self_ir": function["self_ir"],
        "direct_children": direct,
        "direct_children_ir": sum(item["inclusive_ir"] for item in direct),
    }


def incoming_edges(parsed: dict[str, Any], target_id: int) -> list[dict[str, Any]]:
    names = parsed["names"]
    result: list[dict[str, Any]] = []
    for function in parsed["functions"].values():
        for edge in function["edges"]:
            if (edge["callee_id"] == target_id and edge["calls"] > 0
                    and edge["inclusive_ir"] > 0):
                result.append({
                    **edge,
                    "caller": display_name(names.get(edge["caller_id"], "")),
                    "callee": display_name(names.get(edge["callee_id"], "")),
                })
    return result


def ancestry_to_owner(parsed: dict[str, Any], owner_id: int) -> list[dict[str, Any]]:
    """Retain a deterministic positive incoming ancestry where one exists."""

    current = owner_id
    seen = {owner_id}
    result: list[dict[str, Any]] = []
    while len(result) < 32:
        candidates = [edge for edge in incoming_edges(parsed, current)
                      if edge["caller_id"] not in seen]
        if not candidates:
            break
        edge = sorted(candidates,
                      key=lambda item: (-item["inclusive_ir"], item["caller_id"]))[0]
        result.append({
            "caller_id": edge["caller_id"],
            "caller": edge["caller"],
            "callee_id": edge["callee_id"],
            "callee": edge["callee"],
            "calls": edge["calls"],
            "inclusive_ir": edge["inclusive_ir"],
        })
        seen.add(edge["caller_id"])
        current = edge["caller_id"]
    result.reverse()
    return result


def dominant_path(parsed: dict[str, Any], owner_id: int, depth: int = 8) -> list[dict[str, Any]]:
    names = parsed["names"]
    functions = parsed["functions"]
    result: list[dict[str, Any]] = []
    current = owner_id
    seen: set[int] = set()
    for _ in range(depth):
        if current in seen or current not in functions:
            break
        seen.add(current)
        view = function_view(functions[current], names)
        result.append({
            "id": view["id"], "name": view["name"],
            "self_ir": view["self_ir"],
            "inclusive_ir": (result[-1]["edge_inclusive_ir"]
                              if result and "edge_inclusive_ir" in result[-1]
                              else None),
            "direct_children_ir": view["direct_children_ir"],
        })
        children = [item for item in view["direct_children"] if item["callee_id"] in functions]
        if not children:
            break
        edge = children[0]
        result[-1]["next_edge"] = edge
        current = edge["callee_id"]
        # The next node's inclusive value is the edge just followed.
        if current in functions:
            # Store this on a private temporary field consumed on next loop.
            functions[current].setdefault("_path_edge_ir", edge["inclusive_ir"])
    # Remove private state and attach edge-derived inclusive values cleanly.
    for index, node in enumerate(result):
        if index == 0:
            node["inclusive_ir"] = None
        else:
            node["inclusive_ir"] = result[index - 1]["next_edge"]["inclusive_ir"]
    for function in functions.values():
        function.pop("_path_edge_ir", None)
    return result


def source_fixture(shape: str) -> tuple[dict[str, Any], dict[str, Any]]:
    fixture_path = OLD_PACKET / "qualification" / f"0-{shape}-capture-before.json"
    seal_path = OLD_PACKET / "seal.json"
    seal = read_json(seal_path)
    files = seal.get("files")
    require(isinstance(files, dict), "0780 seal file map is missing")
    rel = f"qualification/0-{shape}-capture-before.json"
    require(files.get(rel) == sha256(fixture_path), f"0780 fixture seal differs: {rel}")
    fixture = read_json(fixture_path)
    sample = fixture.get("samples", [None])[0]
    require(isinstance(sample, dict), f"0780 fixture sample is missing: {shape}")
    return fixture, sample


def report_identity(path: Path, shape: str, fixture: dict[str, Any], sample: dict[str, Any],
                    expected_allocator_binary: str) -> dict[str, Any]:
    label = relative(path)
    report = read_json(path)
    require(report.get("schema") == "litchi.pptx.capture-profile-probe.v1",
            f"{label}: profile schema changed")
    require(report.get("tool") == "pptx-capture-probe-0784", f"{label}: tool changed")
    require(report.get("mode") == "capture" and report.get("shape") == shape,
            f"{label}: mode/shape changed")
    require(report.get("timing_scope") == "Package::opened_presentation only",
            f"{label}: timing scope changed")
    require(report.get("marker") == MARKER, f"{label}: marker changed")
    require(report.get("source") == fixture.get("source"), f"{label}: source identity differs")
    require((report.get("slides"), report.get("shapes_per_slide")) == DIMENSIONS[shape],
            f"{label}: shape dimensions changed")
    require(report.get("warmup") == 0 and report.get("samples_requested") == 1,
            f"{label}: profile sample policy changed")
    samples = report.get("samples")
    require(isinstance(samples, list) and len(samples) == 1, f"{label}: profile sample count changed")
    allocator = report.get("allocator")
    require(allocator == {"binary": expected_allocator_binary,
                          "allocator": "Rust system allocator",
                          "instrumentation": "none", "counter_revision": None},
            f"{label}: allocator identity changed")
    identities = []
    expected_output = sample.get("output")
    expected_verification = sample.get("verification")
    for index, item in enumerate(samples):
        require(item.get("index") == index, f"{label}: sample indexes are not contiguous")
        require(type(item.get("elapsed_ns")) is int and item["elapsed_ns"] > 0,
                f"{label}: sample elapsed time is invalid")
        require(item.get("source_sha256") == fixture["source"]["sha256"],
                f"{label}: sample source identity differs")
        require(item.get("output") == expected_output, f"{label}: output identity differs")
        require(item.get("verification") == expected_verification,
                f"{label}: semantic verification differs")
        metrics = item.get("metrics")
        require(metrics == {
            "elapsed_ns": item["elapsed_ns"],
            "slides": DIMENSIONS[shape][0],
            "shapes_per_slide": DIMENSIONS[shape][1],
            "captured_slides": DIMENSIONS[shape][0],
            "captured_shapes_per_slide": DIMENSIONS[shape][1],
        }, f"{label}: sample metrics differ")
        identities.append({"source": item["source_sha256"],
                           "output": item["output"],
                           "verification": item["verification"]})
    return {"source": report["source"], "output": expected_output,
            "verification": expected_verification, "sample_count": len(samples),
            "identities": identities}


def validate_native_identity(block: int, shape: str, profile_identity: dict[str, Any],
                             fixture: dict[str, Any], sample: dict[str, Any]) -> dict[str, Any]:
    path = HERE / "native" / f"{block}-{shape}-profile.json"
    require(path.is_file() and not path.is_symlink(), f"missing native parity report: {path.name}")
    report = read_json(path)
    require(report.get("schema") == "litchi.pptx.capture-profile-probe.v1",
            f"{relative(path)}: native schema changed")
    require(report.get("tool") == "pptx-capture-probe-0784"
            and report.get("mode") == "capture" and report.get("shape") == shape,
            f"{relative(path)}: native identity changed")
    require(report.get("source") == fixture["source"], f"{relative(path)}: native source differs")
    native_samples = report.get("samples")
    require(isinstance(native_samples, list) and len(native_samples) == 30,
            f"{relative(path)}: native sample count changed")
    require(report.get("warmup") == 3 and report.get("samples_requested") == 30,
            f"{relative(path)}: native sample policy changed")
    expected_output = sample["output"]
    expected_verification = sample["verification"]
    require(profile_identity["output"] == expected_output
            and profile_identity["verification"] == expected_verification,
            f"{relative(path)}: profile identity differs from native oracle")
    for index, item in enumerate(native_samples):
        require(item.get("index") == index, f"{relative(path)}: native sample indexes changed")
        require(item.get("source_sha256") == fixture["source"]["sha256"]
                and item.get("output") == expected_output
                and item.get("verification") == expected_verification,
                f"{relative(path)}: native output/semantic identity differs")
    return {"path": relative(path), "sha256": sha256(path), "samples": len(native_samples),
            "identity_matches": True}


def validate_raw_pair(
    numbered: Path, terminal: Path, shape: str, receipt: dict[str, Any], parsed_build: dict[str, Any]
) -> dict[str, Any]:
    number = parse_raw(numbered)
    final = parse_raw(terminal, allow_empty=True)
    label = relative(numbered)
    require(number["header"]["part"] == 1, f"{label}: numbered part changed")
    require(number["header"]["trigger"] == "--dump-after=" + OWNER,
            f"{label}: collection trigger changed")
    unexpected_dump = numbered.with_name(numbered.name.rsplit(".1", 1)[0] + ".2")
    require(not unexpected_dump.exists(),
            f"{label}: unexpected additional numbered dump")
    require(number["header"]["summary_ir"] > 0, f"{label}: scoped summary is zero")
    require(number["header"]["totals_ir"] == number["header"]["summary_ir"],
            f"{label}: totals do not match summary")
    require(final["header"]["part"] == 2 and final["header"]["trigger"] == "Program termination",
            f"{relative(terminal)}: termination dump changed")
    require(final["header"]["summary_ir"] == 0 and final["header"]["totals_ir"] == 0,
            f"{relative(terminal)}: termination dump is not zero Ir")
    require(number["header"]["command"] and final["header"]["command"],
            f"{label}: raw command is missing")

    owner_ids = [fid for fid, function in number["functions"].items()
                 if function["name"] == OWNER]
    require(len(owner_ids) == 1,
            f"{label}: exact owner name resolved to {owner_ids}")
    owner_id = owner_ids[0]
    incoming = incoming_edges(number, owner_id)
    require(len(incoming) == 1 and incoming[0]["calls"] == 1,
            f"{label}: owner incoming positive call is not exactly one")
    require(incoming[0]["inclusive_ir"] == number["header"]["summary_ir"],
            f"{label}: owner incoming Ir differs from summary")
    owner = number["functions"][owner_id]
    view = function_view(owner, number["names"])
    direct = view["direct_children"]
    require(view["self_ir"] + view["direct_children_ir"] == number["header"]["summary_ir"],
            f"{label}: self plus immediate children does not reconstruct summary")
    expected_child = [item for item in direct if item["callee"] == CHILD_EXPECTED]
    require(len(expected_child) == 1, f"{label}: expected direct capture child is missing")
    child_id = expected_child[0]["callee_id"]
    child_function = number["functions"].get(child_id)
    require(child_function is not None, f"{label}: direct child function record is missing")
    child_view = function_view(child_function, number["names"])
    nested = [item for item in child_view["direct_children"] if item["callee"] == NESTED_EXPECTED]
    require(len(nested) == 1, f"{label}: expected nested capture_internal child is missing")
    functions = [function_view(function, number["names"])
                 for function in number["functions"].values()]
    all_function_self_ir = sum(function["self_ir"] for function in functions)
    require(all_function_self_ir == number["header"]["summary_ir"],
            f"{label}: all parsed function self Ir does not reconstruct summary")
    functions.sort(key=lambda item: (-item["self_ir"], item["name"], item["id"]))
    top_self = [item for item in functions if item["self_ir"] > 0][:20]
    return {
        "file": label,
        "sha256": number["sha256"],
        "bytes": number["bytes"],
        "termination": {"file": relative(terminal), "sha256": final["sha256"],
                         "bytes": final["bytes"], "summary_ir": 0},
        "summary_ir": number["header"]["summary_ir"],
        "owner": {
            "id": owner_id,
            "name": OWNER,
            "incoming": incoming[0],
            "self_ir": view["self_ir"],
            "direct_children_ir": view["direct_children_ir"],
            "direct_children": direct,
            "partition_equation": "self_ir + sum(immediate child inclusive_ir) = owner inclusive_ir",
            "partition_disjoint": True,
            "nested_inclusive_rows_excluded": True,
        },
        "dominant_child": expected_child[0],
        "capture_internal": nested[0],
        "ancestry_to_owner": ancestry_to_owner(number, owner_id),
        "dominant_inclusive_path": dominant_path(number, owner_id),
        "top_self_functions": top_self,
        "parser": number["statistics"],
        "all_function_self_ir": all_function_self_ir,
        "validation": {
            "exact_trigger": True,
            "summary_nonzero": True,
            "termination_zero_ir": True,
            "exactly_one_positive_owner_call": True,
            "owner_inclusive_matches_summary": True,
            "self_plus_immediate_children_equals_summary": True,
            "all_function_self_ir_equals_summary": True,
            "expected_direct_child_present": True,
            "expected_nested_child_present": True,
            "compressed_names_resolved": number["statistics"]["compressed_name_declarations"] > 0,
            "relative_or_wildcard_positions_parsed": any(
                number["statistics"]["position_kinds"].get(kind, 0) > 0
                for kind in ("relative", "wildcard")
            ),
        },
    }


def analyze() -> dict[str, Any]:
    plan = validate_plan()
    build = validate_build(plan)
    complete_path = HERE / "profiles" / "complete.json"
    complete = read_json(complete_path)
    require(complete == {
        "build_sha256": build["manifest_sha256"],
        "plan_sha256": sha256(HERE / "plan.json"),
        "processes": 6,
    }, "profile completion receipt changed")
    receipts_path = HERE / "profiles" / "receipts.json"
    rows = read_json(receipts_path)
    require(isinstance(rows, list) and len(rows) == 6, "profile receipt cardinality changed")
    expected_jobs = [(block, shape) for block, order in enumerate(PROFILE_ORDER) for shape in order]
    profiles: list[dict[str, Any]] = []
    for row, (block, shape) in zip(rows, expected_jobs):
        require(isinstance(row, dict), f"profile receipt {block}-{shape} is malformed")
        receipt = validate_profile_receipt(row, plan, build, block, shape)
        paths = receipt["paths"]
        fixture, fixture_sample = source_fixture(shape)
        profile_identity = report_identity(paths[f"{block}-{shape}.json"], shape, fixture,
                                           fixture_sample, "profile")
        native_identity = validate_native_identity(block, shape, profile_identity,
                                                   fixture, fixture_sample)
        raw = validate_raw_pair(paths[f"{block}-{shape}.callgrind.1"],
                                paths[f"{block}-{shape}.callgrind"], shape, row, build)
        raw["block"] = block
        raw["shape"] = shape
        raw["report"] = relative(paths[f"{block}-{shape}.json"])
        raw["report_sha256"] = sha256(paths[f"{block}-{shape}.json"])
        raw["output_identity"] = profile_identity
        raw["native_identity"] = native_identity
        profiles.append(raw)
    require([item["shape"] for item in profiles] == list(SHAPES) + ["large", "medium", "tiny"],
            "profile result order changed")
    self_values = [item["owner"]["self_ir"] for item in profiles]
    inclusive_values = [item["summary_ir"] for item in profiles]
    return {
        "schema": "litchi-0784-callgrind-profile-analysis-v1",
        "packet": "change-0784",
        "plan": {"path": relative(HERE / "plan.json"), "sha256": sha256(HERE / "plan.json"),
                 "schema": plan["schema"], "owner": OWNER, "cpu": plan["cpu"]},
        "build": build["custody"],
        "profile_count": len(profiles),
        "profiles": profiles,
        "summary": {
            "positive_ir_profiles": len(profiles),
            "inclusive_ir_min": min(inclusive_values),
            "inclusive_ir_max": max(inclusive_values),
            "wrapper_self_ir_values": self_values,
            "all_exact_scope_checks_pass": True,
            "all_output_fixture_checks_pass": True,
            "all_native_parity_checks_pass": True,
        },
        "claims": [
            "Callgrind Ir is guest instruction attribution for the wrapper-selected capture region.",
            "Immediate child inclusive Ir rows are a disjoint partition only when combined with wrapper self Ir; nested inclusive rows are diagnostic and are not added.",
            "The native reports and 0780 qualification fixtures provide output and semantic identity checks, not a native latency or historical-causality claim.",
        ],
    }


def markdown(report: dict[str, Any]) -> str:
    lines = [
        "# 0784 Callgrind profile analysis",
        "",
        "This is an offline replay of the six retained Callgrind captures. Ir is",
        "guest instruction attribution for the exact `capture_region_0784` wrapper;",
        "it is not native latency or a production performance claim.",
        "",
        "## Scoped totals",
        "",
        "| Pass | Shape | Summary / owner inclusive Ir | Wrapper self Ir | Immediate child Ir | Dominant child | Dominant child Ir |",
        "| ---: | --- | ---: | ---: | ---: | --- | ---: |",
    ]
    for item in report["profiles"]:
        owner = item["owner"]
        child = item["dominant_child"]
        lines.append(
            f"| {item['block']} | {item['shape']} | {item['summary_ir']:,} | "
            f"{owner['self_ir']:,} | {owner['direct_children_ir']:,} | "
            f"`{child['callee']}` | {child['inclusive_ir']:,} |"
        )
    lines += [
        "",
        "The partition checked for every row is `wrapper self Ir + immediate-child",
        "inclusive Ir = owner inclusive Ir`. Descendant inclusive rows overlap and",
        "are excluded from that sum.",
        "",
        "## Dominant inclusive paths",
        "",
    ]
    for item in report["profiles"]:
        path = item["dominant_inclusive_path"]
        rendered = " -> ".join(
            f"`{node['name']}` ({node['self_ir']:,} self)" for node in path
        )
        lines.append(f"- **{item['block']}-{item['shape']}**: {rendered}")
    lines += [
        "",
        "## Top self Ir rows",
        "",
    ]
    for item in report["profiles"]:
        lines += [f"### {item['block']}-{item['shape']}", "",
                  "| Function | Self Ir | Immediate child Ir |", "| --- | ---: | ---: |"]
        for function in item["top_self_functions"][:10]:
            lines.append(f"| `{function['name']}` | {function['self_ir']:,} | "
                         f"{function['direct_children_ir']:,} |")
        lines.append("")
    lines += [
        "Raw `.callgrind.1` hashes, zero-summary termination hashes, receipt",
        "artifacts, native report parity, and 0780 capture fixture identities are",
        "bound in `profile-analysis.json`.",
        "",
    ]
    return "\n".join(lines)


def write_or_check(path: Path, content: str, check: bool) -> None:
    if path.exists():
        require(path.is_file() and not path.is_symlink(), f"output is not a regular file: {path}")
        require(path.read_text(encoding="utf-8") == content,
                f"replayed output differs: {relative(path)}")
    else:
        require(not check, f"missing expected output: {relative(path)}")
        path.write_text(content, encoding="utf-8")


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--check", action="store_true",
                        help="replay and compare existing JSON/Markdown outputs")
    args = parser.parse_args()
    try:
        report = analyze()
        encoded = json.dumps(report, indent=2, sort_keys=True) + "\n"
        rendered = markdown(report)
        write_or_check(ANALYSIS_JSON, encoded, args.check)
        write_or_check(ANALYSIS_MD, rendered, args.check)
        print(f"0784 profile analysis {'replay' if args.check else 'write'} PASS", flush=True)
        return 0
    except (EvidenceError, OSError, KeyError, TypeError, ValueError) as error:
        print(f"profile_analysis.py: error: {error}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
