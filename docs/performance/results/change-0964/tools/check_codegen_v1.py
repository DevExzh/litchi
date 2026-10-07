#!/usr/bin/env python3
"""Read-only, bounded code-generation comparison for the 0964 candidate.

The tool compares only the two named validation callers and the direct calls
visible in their exact ELF address ranges.  A streaming whole-image objdump
collects the small set of global NsReader edges without retaining its output.
It never runs either executable, builds source, or makes a performance claim.
The admission check covers disappearance of the two wrapper edges.  Resolver
edge counts are retained as raw observations: a missing named call can mean
inlining and is not evidence that namespace work was removed.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import subprocess
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Iterable


SCHEMA = "litchi.performance.0964.codegen-check.v1"

FUNCTIONS: tuple[tuple[str, str], ...] = (
    (
        "worksheet_xml_and_parse_source",
        "litchi_xlsx::cell_values::validation::worksheet_xml_and_parse_source",
    ),
    (
        "validate_xml",
        "litchi_xlsx::cell_values::validation::validate_xml",
    ),
)

EDGE_KINDS = (
    "process_event",
    "resolve_event",
    "resolver_push",
    "resolver_resolve_prefix",
)

NM_LINE = re.compile(r"^\s*([0-9a-fA-F]+)\s+([0-9a-fA-F]+)\s+(\S)\s+(.+?)\s*$")
ADDR_LINE = re.compile(r"^\s*([0-9a-fA-F]+):")
# Rust demangled function names can contain nested generic brackets.
HEADER_LINE = re.compile(r"^\s*([0-9a-fA-F]+)\s+<(.+)>:\s*$")


class CheckError(RuntimeError):
    pass


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def describe_binary(path: Path) -> dict[str, Any]:
    if not path.is_file():
        raise CheckError(f"binary is not a regular file: {path}")
    if not os.access(path, os.X_OK):
        raise CheckError(f"binary is not executable: {path}")

    build_id: str | None = None
    note = subprocess.run(
        ["readelf", "-n", str(path)],
        check=False,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
    )
    if note.returncode == 0:
        for line in note.stdout.splitlines():
            if "Build ID:" in line:
                build_id = line.split("Build ID:", 1)[1].strip()
                break

    stat = path.stat()
    return {
        "path": str(path),
        "bytes": stat.st_size,
        "mode": stat.st_mode & 0o777,
        "sha256": sha256_file(path),
        "build_id": build_id,
    }


def run_text(argv: list[str]) -> str:
    proc = subprocess.run(
        argv,
        check=False,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        errors="replace",
    )
    if proc.returncode != 0:
        detail = proc.stderr.strip().splitlines()[-1:] or ["no diagnostic"]
        raise CheckError(f"{argv[0]} failed rc={proc.returncode}: {detail[0]}")
    return proc.stdout


def parse_nm(path: Path) -> list[dict[str, Any]]:
    records: list[dict[str, Any]] = []
    for line in run_text(
        ["nm", "-anS", "--demangle", "--defined-only", str(path)]
    ).splitlines():
        match = NM_LINE.match(line)
        if not match:
            continue
        address, size, symbol_type, name = match.groups()
        records.append(
            {
                "address": int(address, 16),
                "size": int(size, 16),
                "type": symbol_type,
                "name": name,
            }
        )
    return records


def select_symbol(
    records: Iterable[dict[str, Any]], label: str, wanted: str
) -> dict[str, Any]:
    exact = [record for record in records if record["name"] == wanted]
    if len(exact) != 1:
        short = wanted.rsplit("::", 1)[-1]
        suffix = [
            record
            for record in records
            if record["name"].endswith("::" + short)
        ]
        if len(suffix) != 1:
            names = [record["name"] for record in suffix[:8]]
            raise CheckError(
                f"{label}: expected one symbol {wanted!r}; "
                f"exact={len(exact)} suffix={len(suffix)} names={names}"
            )
        exact = suffix
    record = dict(exact[0])
    if record["size"] <= 0:
        raise CheckError(f"{label}: symbol has no usable size: {record}")
    record["requested_name"] = wanted
    return record


def edge_kind(target: str) -> str | None:
    # objdump may print a target as symbol+0xN or with a PLT suffix.  Keep
    # matching anchored to the fully qualified method suffixes.
    base = target.split("+0x", 1)[0].removesuffix("@plt")
    if "quick_xml::reader::ns_reader::NsReader" in base and base.endswith(
        "::process_event"
    ):
        return "process_event"
    if "quick_xml::name::NamespaceResolver::resolve_event" in base:
        return "resolve_event"
    if "quick_xml::name::NamespaceResolver::push" in base:
        return "resolver_push"
    if "quick_xml::name::NamespaceResolver::resolve_prefix" in base:
        return "resolver_resolve_prefix"
    return None


def parse_call_line(line: str) -> tuple[int, str, str] | None:
    address_match = ADDR_LINE.match(line)
    if not address_match or not re.search(r"\bcall(?:q)?\b", line):
        return None
    # Demangled Rust generic names contain nested angle brackets, e.g.
    # ``NsReader<R>::process_event``.  The target is the final angle-bracket
    # pair on an objdump call line, so use a greedy match here.
    target_match = re.search(r"<(.+)>", line)
    if not target_match:
        return None
    target = target_match.group(1)
    kind = edge_kind(target)
    if kind is None:
        return None
    return int(address_match.group(1), 16), kind, line.strip()


def disassemble_range(path: Path, symbol: dict[str, Any]) -> dict[str, Any]:
    start = symbol["address"]
    stop = start + symbol["size"]
    text = run_text(
        [
            "objdump",
            "-drwC",
            "--no-show-raw-insn",
            f"--start-address=0x{start:x}",
            f"--stop-address=0x{stop:x}",
            str(path),
        ]
    )
    calls: dict[str, list[dict[str, Any]]] = {kind: [] for kind in EDGE_KINDS}
    for line in text.splitlines():
        parsed = parse_call_line(line)
        if parsed is None:
            continue
        address, kind, instruction = parsed
        calls[kind].append(
            {
                "address": f"0x{address:x}",
                "offset": f"0x{address - start:x}",
                "instruction": instruction,
            }
        )
    return {
        "start": f"0x{start:x}",
        "stop_exclusive": f"0x{stop:x}",
        "size": symbol["size"],
        "symbol": symbol["name"],
        "calls": calls,
        "call_counts": {kind: len(calls[kind]) for kind in EDGE_KINDS},
    }


def scan_global_edges(path: Path) -> dict[str, Any]:
    """Scan objdump through a pipe; retain only four qualified edge kinds."""
    process = subprocess.Popen(
        ["objdump", "-drwC", "--no-show-raw-insn", str(path)],
        stdout=subprocess.PIPE,
        stderr=subprocess.DEVNULL,
        text=True,
        errors="replace",
    )
    assert process.stdout is not None
    edges: dict[str, list[dict[str, Any]]] = {kind: [] for kind in EDGE_KINDS}
    current_address: int | None = None
    current_name: str | None = None
    for line in process.stdout:
        header = HEADER_LINE.match(line)
        if header:
            current_address = int(header.group(1), 16)
            current_name = header.group(2)
            continue
        parsed = parse_call_line(line)
        if parsed is None:
            continue
        address, kind, instruction = parsed
        edges[kind].append(
            {
                "address": f"0x{address:x}",
                "caller_address": (
                    f"0x{current_address:x}" if current_address is not None else None
                ),
                "caller": current_name,
                "instruction": instruction,
            }
        )
    returncode = process.wait()
    if returncode != 0:
        raise CheckError(f"objdump global scan failed rc={returncode}")
    return {
        "edge_counts": {kind: len(edges[kind]) for kind in EDGE_KINDS},
        "edges": edges,
    }


def mark_changed_callers(
    scan: dict[str, Any], ranges: dict[str, dict[str, Any]]
) -> dict[str, Any]:
    by_start = {
        int(record["start"], 16): label for label, record in ranges.items()
    }
    for edges in scan["edges"].values():
        for edge in edges:
            caller_address = edge.get("caller_address")
            label = by_start.get(int(caller_address, 16)) if caller_address else None
            edge["changed_caller"] = label
    scan["other_than_changed_callers"] = {
        kind: [edge for edge in scan["edges"][kind] if edge["changed_caller"] is None]
        for kind in EDGE_KINDS
    }
    scan["other_than_changed_counts"] = {
        kind: len(scan["other_than_changed_callers"][kind]) for kind in EDGE_KINDS
    }
    return scan


def inspect_binary(path: Path, global_scan: bool) -> dict[str, Any]:
    records = parse_nm(path)
    selected: dict[str, dict[str, Any]] = {}
    ranges: dict[str, dict[str, Any]] = {}
    for label, wanted in FUNCTIONS:
        selected[label] = select_symbol(records, label, wanted)
        ranges[label] = disassemble_range(path, selected[label])
    symbol_presence = {
        kind: [
            record["name"]
            for record in records
            if edge_kind(record["name"]) == kind
        ]
        for kind in EDGE_KINDS
    }
    result: dict[str, Any] = {
        "functions": ranges,
        "reference_symbol_presence": symbol_presence,
    }
    if global_scan:
        result["global_edges"] = mark_changed_callers(
            scan_global_edges(path), ranges
        )
    return result


def compare_side_edges(
    baseline: dict[str, Any], candidate: dict[str, Any]
) -> dict[str, Any]:
    checks: dict[str, Any] = {}
    for label, _ in FUNCTIONS:
        base_counts = baseline["functions"][label]["call_counts"]
        cand_counts = candidate["functions"][label]["call_counts"]
        checks[label] = {
            "baseline": base_counts,
            "candidate": cand_counts,
            "delta_candidate_minus_baseline": {
                kind: cand_counts[kind] - base_counts[kind] for kind in EDGE_KINDS
            },
            "direct_process_event_disappeared": (
                base_counts["process_event"] > 0
                and cand_counts["process_event"] == 0
            ),
            "direct_resolve_event_disappeared": (
                base_counts["resolve_event"] > 0
                and cand_counts["resolve_event"] == 0
            ),
            "resolver_push_named_edge": cand_counts["resolver_push"] > 0,
            "resolver_resolve_prefix_named_edge": cand_counts[
                "resolver_resolve_prefix"
            ]
            > 0,
            "resolver_push_assembly_observation": (
                "observed_named_call"
                if cand_counts["resolver_push"] > 0
                else "not_observed_may_be_inlined"
            ),
            "resolver_resolve_prefix_assembly_observation": (
                "observed_named_call"
                if cand_counts["resolver_resolve_prefix"] > 0
                else "not_observed_may_be_inlined"
            ),
        }
    return checks


def main(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--baseline", required=True, type=Path)
    parser.add_argument("--candidate", required=True, type=Path)
    parser.add_argument("--output", required=True, type=Path)
    parser.add_argument(
        "--no-global-scan",
        action="store_true",
        help="only inspect the two exact function ranges",
    )
    args = parser.parse_args(argv)
    if args.output.exists():
        raise CheckError(f"refusing to overwrite existing output: {args.output}")

    baseline_desc = describe_binary(args.baseline)
    candidate_desc = describe_binary(args.candidate)
    baseline = inspect_binary(args.baseline, not args.no_global_scan)
    candidate = inspect_binary(args.candidate, not args.no_global_scan)
    checks = compare_side_edges(baseline, candidate)
    required_checks = [
        value[check]
        for value in checks.values()
        for check in (
            "direct_process_event_disappeared",
            "direct_resolve_event_disappeared",
        )
    ]
    status = "passed" if all(required_checks) else "failed"

    result: dict[str, Any] = {
        "schema": SCHEMA,
        "status": status,
        "generated_at_utc": datetime.now(timezone.utc).isoformat(),
        "tool": {
            "path": str(Path(__file__).resolve()),
            "sha256": sha256_file(Path(__file__).resolve()),
            "argv": [str(item) for item in ["python3", *argv]],
        },
        "scope": {
            "read_only": True,
            "runs_workload": False,
            "builds_source": False,
            "performance_claim": False,
            "disassembly_scope": "two exact nm function ranges plus streaming global edge scan",
        },
        "inputs": {"baseline_0957_ordinary": baseline_desc, "candidate_0964": candidate_desc},
        "observations": {
            "baseline": baseline,
            "candidate": candidate,
            "changed_caller_edge_checks": checks,
        },
        "interpretation": {
            "direct_process_event_and_resolve_event_edges": (
                "candidate must have zero direct calls in both changed callers"
            ),
            "resolver_push_and_prefix": (
                "raw named-call counts are reported; missing named calls are inconclusive because the compiler may inline them"
            ),
            "global_other_nsreader_calls": (
                "reported separately because unrelated NsReader users are expected"
            ),
            "namespace_semantic_work": (
                "requires separate source review; this assembly check does not infer removal from named-call absence"
            ),
        },
    }
    args.output.parent.mkdir(parents=True, exist_ok=True)
    with args.output.open("x", encoding="utf-8") as handle:
        json.dump(result, handle, indent=2, sort_keys=True)
        handle.write("\n")
    print(json.dumps({"output": str(args.output), "status": status, "schema": SCHEMA}))
    return 0 if status == "passed" else 1


if __name__ == "__main__":
    try:
        raise SystemExit(main(sys.argv[1:]))
    except CheckError as error:
        print(f"check_codegen_v1: {error}", file=sys.stderr)
        raise SystemExit(2)
