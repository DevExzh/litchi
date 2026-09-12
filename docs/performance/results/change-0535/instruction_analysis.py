#!/usr/bin/env python3
"""Validate and attribute 0535 instruction-position Callgrind evidence.

The report is intentionally diagnostic.  It binds each exclusive
SectorChainScratch::collect_exact Ir cost to an exact instruction in the
retained baseline binary and keeps call/jump metadata separate from the
instruction costs.  It makes no operation-local timing claim.
"""
from __future__ import annotations

import argparse
import collections
import hashlib
import json
import re
import sys
from pathlib import Path
from typing import Any, Iterable


HERE = Path(__file__).resolve().parent
BASELINE = HERE / "baseline"
PLAN_PATH = HERE / "plan.json"
ASSEMBLY_INDEX_PATH = BASELINE / "assembly-index.json"
BINARY_META_PATH = BASELINE / "binary-normal.json"
SOURCE_MANIFEST_PATH = BASELINE / "source-manifest.json"
INSPECT_SCRIPT_PATH = HERE / "inspect_assembly.py"

COLLECTOR_MARKERS = ("SectorChainScratch", "collect_exact")
INSERT_MARKERS = ("CheckedBitSet", "insert")
COLLECTOR_PARENT_MARKER = "validate_stream_allocations"
FUNCTION_RE = re.compile(r"^(fn|cfn)=\((\d+)\)(?:\s+(.*))?$")
OBJECT_RE = re.compile(r"^(ob|cob)=\((\d+)\)(?:\s+(.*))?$")
CALLS_RE = re.compile(r"^calls=([0-9,]+)")
JUMP_RE = re.compile(r"^(jump|jcnd)=(.*)$")
SUMMARY_RE = re.compile(r"^summary:\s*(.*)$")
PART_RE = re.compile(r"^part:\s*(\d+)\s*$")
TRIGGER_RE = re.compile(r"^desc:\s*Trigger:\s*(.*)$")
COST_PREFIX_RE = re.compile(r"^[0-9A-Fa-fxX+\-*]")
HEADER_RE = re.compile(r"^\s*([0-9A-Fa-f]+)\s+<([^>]+)>:\s*$")
INSTRUCTION_RE = re.compile(r"^\s*([0-9A-Fa-f]+):\s+(.*)$")
CONTINUATION_RE = re.compile(r"^\s+((?:[0-9A-Fa-f]{2}\s*)+)$")


class EvidenceError(ValueError):
    """Raised when the frozen evidence is incomplete or inconsistent."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise EvidenceError(message)


def sha256(path: Path) -> str:
    try:
        data = path.read_bytes()
    except OSError as exc:
        raise EvidenceError(f"cannot read {path}: {exc}") from exc
    return hashlib.sha256(data).hexdigest()


def read_json(path: Path) -> dict[str, Any]:
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise EvidenceError(f"cannot read JSON {path}: {exc}") from exc
    require(isinstance(value, dict), f"{path} must contain a JSON object")
    return value


def relpath(path: Path) -> str:
    try:
        return path.relative_to(HERE).as_posix()
    except ValueError:
        return str(path)


def parse_decimal(token: str, what: str) -> int:
    try:
        return int(token.replace(",", ""), 10)
    except ValueError as exc:
        raise EvidenceError(f"invalid {what}: {token!r}") from exc


def parse_position(token: str) -> int:
    # --compress-pos=no is a contract for this run.  Relative positions are
    # rejected rather than being silently interpreted against the wrong base.
    require(token and token[0] not in "+-*", f"compressed instruction position {token!r}")
    try:
        if token.lower().startswith("0x"):
            value = int(token, 16)
        else:
            require(token.isdigit(), f"non-absolute instruction position {token!r}")
            value = int(token, 10)
    except ValueError as exc:
        raise EvidenceError(f"invalid instruction position {token!r}") from exc
    require(value >= 0, f"negative instruction position {token!r}")
    return value


def has_markers(name: str, markers: Iterable[str]) -> bool:
    return all(marker in name for marker in markers)


def caller_matches(actual: str, expected: str) -> bool:
    return actual == expected or actual.endswith("::" + expected)


def _name(value: str | None) -> str:
    return value or ""


def _function(parsed: dict[str, Any], function_id: int) -> dict[str, Any]:
    function = parsed["functions"].get(function_id)
    require(function is not None, f"missing function id {function_id}")
    return function


def parse_cost_line(line: str, line_number: int) -> tuple[int, int]:
    fields = line.split()
    require(len(fields) == 2, f"line {line_number}: expected instr Ir cost, got {line!r}")
    position = parse_position(fields[0])
    cost = parse_decimal(fields[1], f"Ir cost on line {line_number}")
    require(cost >= 0, f"negative Ir cost on line {line_number}")
    return position, cost


def _new_function(function_id: int, name: str, object_id: int | None,
                  object_name: str | None) -> dict[str, Any]:
    return {
        "id": function_id,
        "name": name,
        "object_id": object_id,
        "object": object_name,
        "self_ir": 0,
        "instruction_ir": collections.defaultdict(int),
        "instruction_lines": collections.defaultdict(int),
        "edges": [],
        "jumps": [],
    }


def parse_raw_profile(path: Path) -> dict[str, Any]:
    """Parse one positions: instr, events: Ir Callgrind dump.

    Only fn self-cost lines contribute to instruction Ir.  cfn and jump/jcnd
    records are retained as separate metadata so they cannot be mistaken for
    operation-local instruction costs.
    """
    try:
        lines = path.read_text(encoding="utf-8").splitlines()
    except OSError as exc:
        raise EvidenceError(f"cannot read raw profile {path}: {exc}") from exc

    events: list[str] | None = None
    positions: list[str] | None = None
    summary: int | None = None
    part: int | None = None
    trigger: str | None = None
    object_names: dict[int, str] = {}
    functions: dict[int, dict[str, Any]] = {}
    current_object_id: int | None = None
    current_object_name: str | None = None
    current_function: int | None = None
    pending_call: dict[str, Any] | None = None
    pending_jump: dict[str, Any] | None = None

    for line_number, line in enumerate(lines, 1):
        stripped = line.strip()
        if not stripped:
            continue
        if stripped.startswith("events:"):
            events = stripped[len("events:"):].split()
            continue
        if stripped.startswith("positions:"):
            positions = stripped[len("positions:"):].split()
            continue
        match = SUMMARY_RE.match(stripped)
        if match:
            fields = match.group(1).split()
            require(fields, f"{path}: empty summary")
            summary = parse_decimal(fields[0], f"summary on line {line_number}")
            continue
        match = PART_RE.match(stripped)
        if match:
            part = int(match.group(1))
            continue
        match = TRIGGER_RE.match(stripped)
        if match:
            trigger = match.group(1)
            continue

        match = OBJECT_RE.match(stripped)
        if match:
            kind, ident, name = match.groups()
            if kind == "ob":
                current_object_id = int(ident)
                current_object_name = name or object_names.get(current_object_id)
                if name:
                    object_names[current_object_id] = name
            elif name:
                # A main executable is commonly introduced first as cob on
                # the synthetic below-main edge, followed later by ob=(id)
                # without a repeated path.  Retain that object identity
                # without changing the owning fn context.
                object_names[int(ident)] = name
            # cob describes the callee object and does not change the owning
            # fn context.  Keeping it out of the attribution graph is
            # deliberate; the cfn record identifies the callee function.
            continue

        match = FUNCTION_RE.match(stripped)
        if match:
            kind, ident, name = match.groups()
            function_id = int(ident)
            if function_id not in functions:
                functions[function_id] = _new_function(
                    function_id, name or "", current_object_id, current_object_name
                )
            function = functions[function_id]
            if name:
                function["name"] = name
            if function["object"] is None:
                function["object_id"] = current_object_id
                function["object"] = current_object_name
            pending_call = None
            pending_jump = None
            if kind == "fn":
                current_function = function_id
            else:
                require(current_function is not None,
                        f"{path}: cfn without a current fn on line {line_number}")
                pending_call = {
                    "callee_id": function_id,
                    "callee_name": function["name"],
                    "calls": None,
                }
            continue

        if stripped.startswith("calls="):
            require(pending_call is not None,
                    f"{path}: calls record without cfn on line {line_number}")
            match = CALLS_RE.match(stripped)
            require(match is not None, f"{path}: malformed calls record on line {line_number}")
            pending_call["calls"] = parse_decimal(match.group(1), "call count")
            continue

        match = JUMP_RE.match(stripped)
        if match:
            require(current_function is not None,
                    f"{path}: jump record without fn on line {line_number}")
            pending_call = None
            pending_jump = {"kind": match.group(1), "spec": match.group(2)}
            continue

        # Callgrind uses fl/fi/fe/cfi/totals and other non-cost records that
        # are irrelevant to this report.  They must not reset the current fn.
        if (
            stripped.startswith(("creator:", "version:", "pid:", "cmd:", "desc:",
                                 "fl=", "fi=", "fe=", "cfi=", "totals:",
                                 "ob=", "cob="))
            or stripped.startswith("#")
        ):
            continue

        if positions is None or events is None or current_function is None:
            continue
        if pending_jump is not None and len(stripped.split()) == 1:
            position = parse_position(stripped)
            functions[current_function]["jumps"].append({
                "kind": pending_jump["kind"],
                "spec": pending_jump["spec"],
                "position": position,
                "ir": None,
            })
            pending_jump = None
            continue
        if not COST_PREFIX_RE.match(stripped):
            continue
        position, cost = parse_cost_line(stripped, line_number)
        function = functions[current_function]
        if pending_call is not None:
            require(pending_call["calls"] is not None,
                    f"{path}: cfn has no calls record on line {line_number}")
            function["edges"].append({
                "callee_id": pending_call["callee_id"],
                "callee_name": pending_call["callee_name"],
                "calls": pending_call["calls"],
                "inclusive_ir": cost,
                "position": position,
            })
            pending_call = None
        elif pending_jump is not None:
            function["jumps"].append({
                "kind": pending_jump["kind"],
                "spec": pending_jump["spec"],
                "position": position,
                "ir": cost,
            })
            pending_jump = None
        else:
            function["self_ir"] += cost
            function["instruction_ir"][position] += cost
            function["instruction_lines"][position] += 1

    require(events == ["Ir"], f"{path}: expected events: Ir, got {events!r}")
    require(positions == ["instr"], f"{path}: expected positions: instr, got {positions!r}")
    require(summary is not None, f"{path}: missing summary")
    require(part is not None, f"{path}: missing part")
    require(trigger is not None, f"{path}: missing trigger")
    require(functions, f"{path}: no function records")
    for function in functions.values():
        require(function["self_ir"] >= 0, f"{path}: negative self Ir")
        require(sum(function["instruction_ir"].values()) == function["self_ir"],
                f"{path}: instruction/self Ir mismatch in {function['name']!r}")
    return {
        "path": path,
        "sha256": sha256(path),
        "events": events,
        "positions": positions,
        "summary": summary,
        "part": part,
        "trigger": trigger,
        "object_names": object_names,
        "functions": functions,
    }


def positive_ancestor_path(parsed: dict[str, Any], target_id: int,
                           ancestor: str, max_depth: int = 8) -> dict[str, Any] | None:
    """Find an exact positive Ir caller path to target_id.

    Call counts may be zero for collection-off call metadata; the positive
    edge predicate is inclusive Ir > 0 with calls >= 0.  This path is used
    only to classify setup/timed ownership and is emitted as evidence.
    """
    queue: collections.deque[tuple[int, list[int], list[dict[str, Any]]]] = (
        collections.deque([(target_id, [target_id], [])])
    )
    seen = {(target_id, 0)}
    while queue:
        child_id, path_ids, path_edges = queue.popleft()
        if len(path_edges) >= max_depth:
            continue
        for parent in parsed["functions"].values():
            for edge in parent["edges"]:
                if edge["callee_id"] != child_id:
                    continue
                if edge["calls"] is None or edge["calls"] < 0:
                    continue
                if edge["inclusive_ir"] <= 0:
                    continue
                candidate_ids = [parent["id"], *path_ids]
                candidate_edges = [{
                    "caller_id": parent["id"],
                    "caller_name": parent["name"],
                    **edge,
                }, *path_edges]
                if caller_matches(parent["name"], ancestor):
                    return {
                        "ancestor": ancestor,
                        "function_ids": candidate_ids,
                        "functions": [parsed["functions"][ident]["name"]
                                      for ident in candidate_ids],
                        "edges": candidate_edges,
                    }
                key = (parent["id"], len(candidate_edges))
                if key not in seen:
                    seen.add(key)
                    queue.append((parent["id"], candidate_ids, candidate_edges))
    return None


def owner_incoming_edges(parsed: dict[str, Any], owner_ids: set[int]) -> list[dict[str, Any]]:
    incoming = []
    for parent in parsed["functions"].values():
        for edge in parent["edges"]:
            if edge["callee_id"] in owner_ids:
                incoming.append({
                    "caller_id": parent["id"],
                    "caller_name": parent["name"],
                    **edge,
                })
    return incoming


def find_owner(parsed: dict[str, Any], owner: str) -> tuple[int, dict[str, Any]]:
    matches = [
        (ident, function) for ident, function in parsed["functions"].items()
        if owner in function["name"]
    ]
    require(matches, f"{parsed['path']}: owner {owner!r} is absent")
    require(len(matches) == 1,
            f"{parsed['path']}: owner {owner!r} has {len(matches)} function ids")
    return matches[0]


def classify_scope(parsed: dict[str, Any], job: dict[str, Any],
                   owner_id: int, owner_function: dict[str, Any]) -> dict[str, Any]:
    incoming = owner_incoming_edges(parsed, {owner_id})
    positive = [
        edge for edge in incoming
        if edge["calls"] is not None
        and edge["calls"] > 0
        and edge["inclusive_ir"] > 0
    ]
    require(len(positive) == 1,
            f"{parsed['path']}: expected one positive incoming owner edge, got {len(positive)}")
    owner_edge = positive[0]
    require(owner_edge["inclusive_ir"] == parsed["summary"],
            f"{parsed['path']}: owner incoming Ir differs from dump summary")
    caller = owner_edge["caller_name"]
    runner_path = positive_ancestor_path(parsed, owner_edge["caller_id"], job["runner"])
    setup_paths = [
        positive_ancestor_path(parsed, owner_edge["caller_id"], setup)
        for setup in job.get("setup_callers", [])
    ]
    setup_paths = [path for path in setup_paths if path is not None]

    runner_direct = caller_matches(caller, job["runner"])
    setup_direct = any(caller_matches(caller, setup)
                       for setup in job.get("setup_callers", []))
    if runner_direct or runner_path is not None:
        require(not setup_direct and not setup_paths,
                f"{parsed['path']}: owner has both timed and setup positive ancestry")
        role = "timed"
        path = runner_path or {
            "ancestor": job["runner"],
            "function_ids": [owner_edge["caller_id"], owner_id],
            "functions": [caller, owner_function["name"]],
            "edges": [owner_edge],
        }
    elif setup_direct or setup_paths:
        role = "setup"
        path = setup_paths[0] if setup_paths else {
            "ancestor": next(
                setup for setup in job["setup_callers"]
                if caller_matches(caller, setup)
            ),
            "function_ids": [owner_edge["caller_id"], owner_id],
            "functions": [caller, owner_function["name"]],
            "edges": [owner_edge],
        }
    else:
        raise EvidenceError(
            f"{parsed['path']}: owner caller {caller!r} has no expected positive ancestry"
        )

    # The search starts at the owner caller so that its result can be checked
    # against the runner/setup graph.  Close the emitted path with the
    # positive caller -> owner edge as well; the resulting path is an exact
    # ancestor-to-owner attribution chain.
    if path["function_ids"][-1] != owner_id:
        path = {
            **path,
            "function_ids": [*path["function_ids"], owner_id],
            "functions": [*path["functions"], owner_function["name"]],
            "edges": [*path["edges"], owner_edge],
        }

    if role == "setup":
        require(job["kind"] == "cfb", f"{parsed['path']}: non-CFB setup dump")
        require(parsed["part"] == 1,
                f"{parsed['path']}: setup ancestry appears in part {parsed['part']}")
    else:
        expected_first = 1 if job["kind"] == "xls" else 2
        require(parsed["part"] >= expected_first,
                f"{parsed['path']}: timed ancestry appears in setup part {parsed['part']}")

    owner_direct_ir = sum(edge["inclusive_ir"] for edge in owner_function["edges"])
    require(owner_function["self_ir"] + owner_direct_ir == owner_edge["inclusive_ir"],
            f"{parsed['path']}: owner parent attribution does not close")
    return {
        "role": role,
        "owner_id": owner_id,
        "owner_name": owner_function["name"],
        "owner_self_ir": owner_function["self_ir"],
        "owner_direct_ir": owner_direct_ir,
        "owner_incoming": owner_edge,
        "positive_path": path,
        "runner_path": runner_path,
        "setup_paths": setup_paths,
    }


def parse_instruction_text(line: str, line_number: int) -> dict[str, Any] | None:
    match = INSTRUCTION_RE.match(line)
    if not match:
        return None
    address = int(match.group(1), 16)
    fields = match.group(2).strip().split()
    byte_tokens: list[str] = []
    while fields and re.fullmatch(r"[0-9A-Fa-f]{2}", fields[0]):
        byte_tokens.append(fields.pop(0))
    if not byte_tokens:
        return None
    if not fields:
        return {
            "address": address,
            "bytes_hex": "".join(byte_tokens).lower(),
            "text": None,
            "continuation": True,
        }
    return {
        "address": address,
        "bytes_hex": "".join(byte_tokens).lower(),
        "text": " ".join(fields),
        "continuation": False,
    }


def parse_objdump(row: dict[str, Any], stdout_path: Path) -> dict[str, Any]:
    try:
        lines = stdout_path.read_text(encoding="utf-8").splitlines()
    except OSError as exc:
        raise EvidenceError(f"cannot read assembly output {stdout_path}: {exc}") from exc
    base = int(str(row["address_hex"]), 16)
    size = int(row["size_bytes"])
    require(size > 0, f"{stdout_path}: non-positive symbol size")
    end = base + size
    header = None
    instructions: dict[int, dict[str, Any]] = {}
    for line_number, line in enumerate(lines, 1):
        match = HEADER_RE.match(line)
        if match:
            header = {"address": int(match.group(1), 16), "symbol": match.group(2)}
            continue
        instruction = parse_instruction_text(line, line_number)
        if instruction is None:
            # GNU objdump wraps long x86 instruction encodings onto a
            # continuation line containing only the remaining bytes.  Attach
            # those bytes to the preceding instruction.
            continuation = CONTINUATION_RE.match(line)
            if continuation and instructions:
                last_address = next(reversed(instructions))
                extra = "".join(continuation.group(1).split()).lower()
                instructions[last_address]["bytes_hex"] += extra
            continue
        if instruction["continuation"]:
            require(instructions, f"{stdout_path}: orphan instruction continuation")
            last_address = next(reversed(instructions))
            require(instruction["address"] == last_address
                    + len(instructions[last_address]["bytes_hex"]) // 2,
                    f"{stdout_path}: discontinuous instruction continuation")
            instructions[last_address]["bytes_hex"] += instruction["bytes_hex"]
            continue
        address = instruction["address"]
        require(base <= address < end,
                f"{stdout_path}: instruction {address:#x} outside symbol range")
        require(address not in instructions,
                f"{stdout_path}: duplicate instruction address {address:#x}")
        instructions[address] = instruction
    require(header is not None, f"{stdout_path}: missing disassembly header")
    require(header["address"] == base,
            f"{stdout_path}: header {header['address']:#x} != nm base {base:#x}")
    require(instructions, f"{stdout_path}: no instructions")
    return {
        "name": row["name"],
        "symbol": row["symbol"],
        "address": base,
        "size_bytes": size,
        "instructions": instructions,
    }


def load_assembly() -> dict[str, Any]:
    index = read_json(ASSEMBLY_INDEX_PATH)
    require(index.get("plan_sha256") == sha256(PLAN_PATH),
            "assembly index is bound to a different plan")
    require(index.get("source_manifest_sha256") == sha256(SOURCE_MANIFEST_PATH),
            "assembly index is bound to a different source manifest")
    require(index.get("script_sha256") == sha256(INSPECT_SCRIPT_PATH),
            "assembly index is bound to a different inspection script")
    rows = index.get("rows")
    require(isinstance(rows, list) and rows, "assembly index has no rows")
    parsed_rows = []
    for row in rows:
        require(isinstance(row, dict), "assembly index row is not an object")
        for key in ("name", "symbol", "address_hex", "size_bytes", "receipt_sha256"):
            require(key in row, f"assembly index row lacks {key}")
        stdout_path = BASELINE / f"{row['name']}.stdout"
        receipt_path = BASELINE / f"{row['name']}.receipt.json"
        require(stdout_path.is_file(), f"missing assembly stdout {stdout_path}")
        require(receipt_path.is_file(), f"missing assembly receipt {receipt_path}")
        require(sha256(receipt_path) == row["receipt_sha256"],
                f"assembly receipt hash mismatch for {row['name']}")
        parsed_rows.append(parse_objdump(row, stdout_path))
    collector_rows = [
        row for row in parsed_rows if has_markers(row["symbol"], COLLECTOR_MARKERS)
    ]
    insert_rows = [
        row for row in parsed_rows if has_markers(row["symbol"], INSERT_MARKERS)
    ]
    require(collector_rows, "assembly index has no SectorChainScratch::collect_exact rows")
    require(insert_rows, "assembly index has no CheckedBitSet::insert rows")
    all_instructions: dict[int, dict[str, Any]] = {}
    for row in parsed_rows:
        for address, instruction in row["instructions"].items():
            previous = all_instructions.get(address)
            require(previous is None or previous["symbol"] == row["symbol"],
                    f"assembly instruction address collision at {address:#x}")
            all_instructions[address] = {
                **instruction,
                "symbol": row["symbol"],
                "function_start": row["address"],
                "function_size": row["size_bytes"],
            }
    collector_instructions: dict[int, dict[str, Any]] = {}
    for row in collector_rows:
        for address, instruction in row["instructions"].items():
            previous = collector_instructions.get(address)
            require(previous is None or previous["symbol"] == row["symbol"],
                    f"collector instruction address collision at {address:#x}")
            collector_instructions[address] = {
                **instruction,
                "symbol": row["symbol"],
                "function_start": row["address"],
                "function_size": row["size_bytes"],
            }
    return {
        "index": index,
        "rows": parsed_rows,
        "collector_rows": collector_rows,
        "insert_rows": insert_rows,
        "all_instructions": all_instructions,
        "collector_instructions": collector_instructions,
    }


def candidate_biases(raw_addresses: set[int],
                     assembly_addresses: set[int]) -> set[int]:
    require(raw_addresses, "no collector instruction addresses were recorded")
    require(assembly_addresses, "assembly has no collector instructions")
    first = min(raw_addresses)
    candidates: set[int] = set()
    for assembly in assembly_addresses:
        bias = first - assembly
        if all((raw - bias) in assembly_addresses for raw in raw_addresses):
            candidates.add(bias)
    return candidates


def find_relocation_bias(parsed_dumps: list[dict[str, Any]],
                         assembly_addresses: set[int]) -> int:
    raw_addresses: set[int] = set()
    for parsed in parsed_dumps:
        for function in parsed["functions"].values():
            if has_markers(function["name"], COLLECTOR_MARKERS):
                raw_addresses.update(function["instruction_ir"])
    candidates = candidate_biases(raw_addresses, assembly_addresses)
    require(len(candidates) == 1,
            f"expected one relocation bias, got {sorted(candidates)}")
    return next(iter(candidates))


def edge_record(edge: dict[str, Any]) -> dict[str, Any]:
    return {
        "callee_id": edge["callee_id"],
        "callee": edge["callee_name"],
        "calls": edge["calls"],
        "inclusive_ir": edge["inclusive_ir"],
        "position_hex": f"0x{edge['position']:x}",
    }


def jump_record(jump: dict[str, Any]) -> dict[str, Any]:
    return {
        "kind": jump["kind"],
        "spec": jump["spec"],
        "position_hex": f"0x{jump['position']:x}",
        "ir": jump["ir"],
    }


def collector_record(parsed: dict[str, Any], scope: dict[str, Any],
                     bias: int, assembly: dict[str, Any],
                     binary_name: str) -> dict[str, Any]:
    functions = [
        function for function in parsed["functions"].values()
        if has_markers(function["name"], COLLECTOR_MARKERS)
    ]
    require(functions, f"{parsed['path']}: collector function is absent")
    functions.sort(key=lambda function: (function["name"], function["id"]))
    records = []
    total_self_ir = 0
    total_instruction_ir = 0
    incoming = []
    for function in functions:
        require(function["object"] is not None
                and Path(function["object"]).name == binary_name,
                f"{parsed['path']}: collector is not owned by {binary_name!r}")
        self_ir = function["self_ir"]
        instruction_ir = sum(function["instruction_ir"].values())
        require(instruction_ir == self_ir,
                f"{parsed['path']}: collector function/self Ir mismatch")
        total_self_ir += self_ir
        total_instruction_ir += instruction_ir
        caller_records = incoming_callers(parsed, function["id"])
        positive_parent_records = [
            record for record in caller_records
            if COLLECTOR_PARENT_MARKER in record["caller"]
            and record["calls"] is not None
            and record["calls"] > 0
            and record["inclusive_ir"] > 0
        ]
        require(positive_parent_records,
                f"{parsed['path']}: collector lacks a positive "
                f"{COLLECTOR_PARENT_MARKER} parent edge")
        instructions = []
        for raw_address in sorted(function["instruction_ir"]):
            normalized = raw_address - bias
            require(normalized >= 0,
                    f"{parsed['path']}: negative normalized instruction address")
            bound = assembly["collector_instructions"].get(normalized)
            require(bound is not None,
                    f"{parsed['path']}: instruction {raw_address:#x} "
                    f"(normalized {normalized:#x}) is absent from objdump")
            instructions.append({
                "address_raw_hex": f"0x{raw_address:x}",
                "address_hex": f"0x{normalized:x}",
                "offset_hex": f"0x{normalized - bound['function_start']:x}",
                "ir": function["instruction_ir"][raw_address],
                "cost_line_count": function["instruction_lines"][raw_address],
                "assembly_symbol": bound["symbol"],
                "bytes_hex": bound["bytes_hex"],
                "assembly": bound["text"],
            })
        records.append({
            "function_id": function["id"],
            "function": function["name"],
            "object": function["object"],
            "self_ir": self_ir,
            "instruction_ir": instructions,
            "direct_callees": [
                edge_record(edge) for edge in sorted(
                    function["edges"],
                    key=lambda edge: (
                        edge["callee_name"], edge["callee_id"],
                        edge["position"], edge["inclusive_ir"],
                    )
                )
            ],
            "jump_metadata": [
                jump_record(jump) for jump in sorted(
                    function["jumps"],
                    key=lambda jump: (jump["kind"], jump["position"], jump["spec"])
                )
            ],
            "incoming_callers": caller_records,
            "positive_parent_callers": positive_parent_records,
            "instruction_ir_sum": instruction_ir,
            "instruction_ir_matches_self": instruction_ir == self_ir,
        })
    require(total_instruction_ir == total_self_ir,
            f"{parsed['path']}: collector instruction Ir does not match function self Ir")
    return {
        "functions": records,
        "function_count": len(records),
        "function_self_ir": total_self_ir,
        "instruction_ir_sum": total_instruction_ir,
        "instruction_ir_matches_function_self": total_instruction_ir == total_self_ir,
        "owner_parent_attribution": {
            "role": scope["role"],
            "owner_id": scope["owner_id"],
            "owner": scope["owner_name"],
            "owner_self_ir": scope["owner_self_ir"],
            "owner_direct_ir": scope["owner_direct_ir"],
            "incoming": edge_record(scope["owner_incoming"]),
            "positive_path": scope["positive_path"],
        },
    }


def incoming_callers(parsed: dict[str, Any], callee_id: int) -> list[dict[str, Any]]:
    records = []
    for parent in parsed["functions"].values():
        for edge in parent["edges"]:
            if edge["callee_id"] == callee_id:
                records.append({
                    "caller_id": parent["id"],
                    "caller": parent["name"],
                    **edge_record(edge),
                })
    records.sort(key=lambda record: (
        record["caller"], record["caller_id"], record["position_hex"],
        record["inclusive_ir"], record["calls"] if record["calls"] is not None else -1,
    ))
    return records


def job_definitions(plan: dict[str, Any]) -> list[dict[str, Any]]:
    profile = plan["profile"]
    repeats = int(profile["repeats"])
    require(repeats == 2, f"expected two profile repeats, got {repeats}")
    jobs = []
    for repeat in range(1, repeats + 1):
        jobs.extend([
            {
                "name": f"profile-r{repeat}-xls-owned",
                "repeat": repeat,
                "kind": "xls",
                "runner": "litchi_perf_baseline::run_xls_owned_source_case",
                "owner": profile["xls_owner"],
                "setup_callers": [],
                "parts": 5,
            },
            {
                "name": f"profile-r{repeat}-cfb-few-large",
                "repeat": repeat,
                "kind": "cfb",
                "runner": "litchi_perf_baseline::run_cfb_open",
                "owner": profile["cfb_owner"],
                "setup_callers": [
                    "litchi_perf_baseline::build_cfb_corpus",
                    "litchi_perf_baseline::run::{{closure}}",
                ],
                "parts": 6,
            },
        ])
    return jobs


def load_dumps(job: dict[str, Any]) -> list[dict[str, Any]]:
    dumps = []
    for number in range(1, job["parts"] + 1):
        path = BASELINE / f"{job['name']}.callgrind.{number}"
        require(path.is_file(), f"missing raw profile {path}")
        parsed = parse_raw_profile(path)
        require(parsed["part"] == number,
                f"{path}: file number {number} has part {parsed['part']}")
        parsed["dump_number"] = number
        dumps.append(parsed)
    return dumps


def aggregate_instruction_ir(dump_records: list[dict[str, Any]],
                             role: str) -> list[dict[str, Any]]:
    totals: dict[str, dict[str, Any]] = {}
    for dump in dump_records:
        if dump["role"] != role:
            continue
        for function in dump["collector"]["functions"]:
            for instruction in function["instruction_ir"]:
                key = instruction["address_hex"]
                row = totals.setdefault(key, {
                    "address_hex": key,
                    "assembly_symbol": instruction["assembly_symbol"],
                    "bytes_hex": instruction["bytes_hex"],
                    "assembly": instruction["assembly"],
                    "ir": 0,
                    "cost_line_count": 0,
                    "dump_count": 0,
                })
                require(row["assembly_symbol"] == instruction["assembly_symbol"],
                        f"assembly symbol changed for {key}")
                row["ir"] += instruction["ir"]
                row["cost_line_count"] += instruction["cost_line_count"]
                row["dump_count"] += 1
    return [totals[key] for key in sorted(totals, key=lambda value: int(value, 16))]


def validate_plan(plan: dict[str, Any]) -> None:
    require(plan.get("status") == "frozen-before-build-and-captures",
            "0535 plan is not the frozen plan")
    profile = plan.get("profile", {})
    require(profile.get("repeats") == 2, "0535 profile repeats must be two")
    require(profile.get("samples") == 5, "0535 timed sample count must be five")
    require(profile.get("warmup") == 0, "0535 profile warmup must be zero")
    require(profile.get("options") == [
        "--dump-instr=yes", "--dump-line=no", "--compress-pos=no",
        "--collect-jumps=yes",
    ], "0535 Callgrind options do not match the frozen contract")
    assembly = plan.get("assembly", {})
    require(assembly.get("owners") == ["SectorChainScratch", "CheckedBitSet"],
            "0535 assembly owners do not match the frozen contract")
    require(assembly.get("required") == ["collect_exact", "insert"],
            "0535 assembly required functions do not match the frozen contract")


def make_report() -> dict[str, Any]:
    plan = read_json(PLAN_PATH)
    validate_plan(plan)
    binary_meta = read_json(BINARY_META_PATH)
    source_manifest_sha = sha256(SOURCE_MANIFEST_PATH)
    require(binary_meta.get("source_manifest_sha256") == source_manifest_sha,
            "binary identity is bound to a different source manifest")
    assembly = load_assembly()
    require(binary_meta.get("sha256") == assembly["index"].get("binary_sha256"),
            "assembly index and binary identity disagree")
    binary_name = Path(str(binary_meta.get("path", ""))).name
    require(binary_name, "binary identity has no executable path")

    jobs_output = []
    all_dump_records = []
    for job in job_definitions(plan):
        parsed_dumps = load_dumps(job)
        bias = find_relocation_bias(parsed_dumps, set(assembly["collector_instructions"]))
        job_dumps = []
        for parsed in parsed_dumps:
            owner_id, owner_function = find_owner(parsed, job["owner"])
            scope = classify_scope(parsed, job, owner_id, owner_function)
            collector = collector_record(parsed, scope, bias, assembly, binary_name)
            record = {
                "number": parsed["dump_number"],
                "part": parsed["part"],
                "role": scope["role"],
                "path": relpath(parsed["path"]),
                "sha256": parsed["sha256"],
                "trigger": parsed["trigger"],
                "summary_ir": parsed["summary"],
                "relocation_bias_hex": f"0x{bias:x}" if bias >= 0 else f"-0x{-bias:x}",
                "owner": {
                    "id": scope["owner_id"],
                    "name": scope["owner_name"],
                    "self_ir": scope["owner_self_ir"],
                    "direct_ir": scope["owner_direct_ir"],
                    "incoming": edge_record(scope["owner_incoming"]),
                    "positive_path": scope["positive_path"],
                },
                "collector": collector,
                "positions": parsed["positions"],
                "events": parsed["events"],
            }
            job_dumps.append(record)
            all_dump_records.append(record)
        require(
            [record["role"] for record in job_dumps].count("timed")
            == (5 if job["kind"] == "xls" else 5),
            f"{job['name']}: expected five timed dumps",
        )
        if job["kind"] == "cfb":
            require(job_dumps[0]["role"] == "setup",
                    f"{job['name']}: first dump must be setup")
            require([record["role"] for record in job_dumps[1:]] == ["timed"] * 5,
                    f"{job['name']}: parts 2-6 must be timed")
        else:
            require([record["role"] for record in job_dumps] == ["timed"] * 5,
                    f"{job['name']}: all dumps must be timed")
        jobs_output.append({
            "name": job["name"],
            "repeat": job["repeat"],
            "kind": job["kind"],
            "runner": job["runner"],
            "owner": job["owner"],
            "dump_count": len(job_dumps),
            "timed_dump_count": sum(record["role"] == "timed" for record in job_dumps),
            "setup_dump_count": sum(record["role"] == "setup" for record in job_dumps),
            "relocation_bias_hex": (
                f"0x{bias:x}" if bias >= 0 else f"-0x{-bias:x}"
            ),
            "dumps": job_dumps,
        })

    require(len(all_dump_records) == 22,
            f"expected 22 profile dumps, got {len(all_dump_records)}")
    require(all(record["collector"]["instruction_ir_matches_function_self"]
                for record in all_dump_records),
            "at least one collector instruction/self Ir check failed")
    return {
        "schema": "litchi-ole2-change-0535-instruction-analysis-v1",
        "scope": (
            "Baseline Callgrind positions:instr attribution for exclusive "
            "SectorChainScratch::collect_exact; call/jump metadata is separate "
            "and is not an operation-local count or latency claim."
        ),
        "performance_claim": None,
        "plan_sha256": sha256(PLAN_PATH),
        "source_manifest_sha256": source_manifest_sha,
        "binary_sha256": binary_meta["sha256"],
        "assembly_index_sha256": sha256(ASSEMBLY_INDEX_PATH),
        "assembly": {
            "rows": len(assembly["rows"]),
            "collector_rows": len(assembly["collector_rows"]),
            "insert_rows": len(assembly["insert_rows"]),
            "collector_instruction_count": len(assembly["collector_instructions"]),
            "required_fragments_present": ["collect_exact", "insert"],
        },
        "jobs": jobs_output,
        "aggregates": {
            "timed": aggregate_instruction_ir(all_dump_records, "timed"),
            "setup": aggregate_instruction_ir(all_dump_records, "setup"),
        },
        "validation": {
            "raw_dump_count": len(all_dump_records),
            "timed_dump_count": sum(record["role"] == "timed"
                                   for record in all_dump_records),
            "setup_dump_count": sum(record["role"] == "setup"
                                    for record in all_dump_records),
            "positions_instr_only": all(record["positions"] == ["instr"]
                                        for record in all_dump_records),
            "events_ir_only": all(record["events"] == ["Ir"]
                                  for record in all_dump_records),
            "single_bias_per_job": all(
                all(dump["relocation_bias_hex"] == job["relocation_bias_hex"]
                    for dump in job["dumps"])
                for job in jobs_output
            ),
            "single_bias_across_profiles": len({
                job["relocation_bias_hex"] for job in jobs_output
            }) == 1,
            "collector_instruction_ir_equals_function_self": all(
                record["collector"]["instruction_ir_matches_function_self"]
                for record in all_dump_records
            ),
            "collector_parent_attributed": all(
                all(function["positive_parent_callers"]
                    for function in record["collector"]["functions"])
                for record in all_dump_records
            ),
            "setup_and_timed_separated": (
                sum(record["role"] == "setup" for record in all_dump_records) == 2
                and sum(record["role"] == "timed" for record in all_dump_records) == 20
            ),
            "direct_callees_separate_from_instruction_ir": all(
                "direct_callees" in function
                and "jump_metadata" in function
                for record in all_dump_records
                for function in record["collector"]["functions"]
            ),
            "no_operation_local_timing_claim": True,
        },
    }


def write_report(report: dict[str, Any], output: Path) -> None:
    data = (json.dumps(report, indent=2, sort_keys=True) + "\n").encode("utf-8")
    if output.exists():
        try:
            existing = output.read_bytes()
        except OSError as exc:
            raise EvidenceError(f"cannot read existing output {output}: {exc}") from exc
        require(existing == data,
                f"refusing to overwrite non-identical report {output}")
        return
    output.parent.mkdir(parents=True, exist_ok=True)
    try:
        output.write_bytes(data)
    except OSError as exc:
        raise EvidenceError(f"cannot write report {output}: {exc}") from exc


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--output", "-o", type=Path, default=HERE / "instruction-analysis.json",
        help="report path (defaults to change-0535/instruction-analysis.json)",
    )
    args = parser.parse_args(argv)
    try:
        report = make_report()
        write_report(report, args.output)
    except EvidenceError as exc:
        print(f"instruction-analysis: error: {exc}", file=sys.stderr)
        return 2
    print(args.output)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
