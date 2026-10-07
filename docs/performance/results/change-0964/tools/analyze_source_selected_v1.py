#!/usr/bin/env python3
"""Bounded source-selected CPU attribution for the retained 0962 wave.

This helper reads the already retained paired perf-script exports.  It uses
the frozen 0962 parser and qualification gates, then restricts the qualified
population to the exact physical ``Snapshot::from_source_selected`` ancestry.
The three requested physical symbols are counted by the deepest physical
frame (an exclusive self population); all other leaves remain in one disjoint
bucket.  Runtime IPs are retained only as evidence alongside the exact
symbol-relative offsets printed by perf.  No source-file coverage claim is
derived from those offsets.

The helper performs no perf invocation, workload launch, build, or source
rewrite.  It runs ``nm``, ``objdump`` and ``addr2line`` read-only against the
sealed executable to bind symbol ranges and instruction text.
"""

from __future__ import annotations

import argparse
import collections
import concurrent.futures
import datetime
import hashlib
import json
import re
import subprocess
import sys
from pathlib import Path
from typing import Any

sys.path.insert(0, str(Path("/home/zhuhe/litchi-performance-0962/tools")))
import analyze_native_cpu_v1 as frozen  # noqa: E402


SCHEMA = "litchi.performance.0964.source-selected-cpu.v1"
ROOT_0962 = Path("/home/zhuhe/litchi-performance-0962")
DEFAULT_WAVE_ROOT = ROOT_0962 / "fp-wave-v1"
DEFAULT_OUTPUT = Path("/home/zhuhe/litchi-performance-0964/source-selected-cpu-receipt-v1.json")
BINARY = ROOT_0962 / "executables/xlsx_managed_compression_benchmark.fp-debug2.v1"
FROZEN_ANALYZER = ROOT_0962 / "tools/analyze_native_cpu_v1.py"
FROZEN_ANALYZER_SHA256 = "f5f7a387f99ba3a587eeb8f27c8f422a9fa696306c2ae47724952b861910bd3d"
BINARY_SHA256 = "843ae35a0f5131a8cba965f917b2a4c5fccb6926ae3d66862e5f78512820fb33"
BINARY_BUILD_ID = "7eb92ba2f0f8f3b248b77d8691e27b80b114fb22"
EXPECTED_SOURCE_SELECTED = {"samples": 7645, "period": 32395631006}
EXPECTED_QUALIFIED = {"samples": 12319, "period": 52852646778}
EVENT = frozen.DEFAULT_EVENT
SOURCE_SYMBOL = "litchi_xlsx::cell_values::snapshot::Snapshot::from_source_selected"
REWRITTEN_SYMBOL = "litchi_xlsx::cell_values::snapshot::Snapshot::from_rewritten_value_source"
TARGETS = {
    "resolve_event": "quick_xml::name::NamespaceResolver::resolve_event",
    "process_event": "quick_xml::reader::ns_reader::NsReader<R>::process_event",
    "worksheet_xml_and_parse_source": "litchi_xlsx::cell_values::validation::worksheet_xml_and_parse_source",
}
OFFSET_RE = re.compile(r"\+0x([0-9a-fA-F]+)$")
DISASM_RE = re.compile(r"^\s*([0-9a-fA-F]+):\s*(.*)$")
MAX_ADDR2LINE_OFFSETS = 16


class AttributionError(RuntimeError):
    """An input identity, qualification, or conservation failure."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise AttributionError(message)


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def descriptor(path: Path) -> dict[str, Any]:
    require(path.is_absolute() and path.is_file() and not path.is_symlink(), f"invalid file: {path}")
    return {"path": str(path), "bytes": path.stat().st_size, "sha256": sha256_file(path)}


def now() -> str:
    return datetime.datetime.now(datetime.timezone.utc).isoformat()


def empty_stats() -> dict[str, int]:
    return {"samples": 0, "period": 0}


def add_stats(target: dict[str, int], sample: frozen.Sample) -> None:
    target["samples"] += 1
    target["period"] += sample.period


def stats(value: dict[str, int], denominator: int) -> dict[str, Any]:
    return {
        "samples": value["samples"],
        "period": value["period"],
        "period_share": value["period"] / denominator if denominator else 0.0,
    }


def exact_symbol(frame: frozen.Frame) -> str:
    return frozen.function_symbol(frame.symbol)


def exact_frames(sample: frozen.Sample, symbol: str) -> list[frozen.Frame]:
    return [
        frame
        for frame in sample.frames
        if exact_symbol(frame) == symbol and frame.dso == str(BINARY)
    ]


def symbol_frames_any_dso(sample: frozen.Sample, symbol: str) -> list[frozen.Frame]:
    return [frame for frame in sample.frames if exact_symbol(frame) == symbol]


def classify_boundary(sample: frozen.Sample) -> tuple[str, dict[str, int]]:
    """Classify exact full physical symbols; no fuzzy namespace matching."""
    source_any = symbol_frames_any_dso(sample, SOURCE_SYMBOL)
    rewritten_any = symbol_frames_any_dso(sample, REWRITTEN_SYMBOL)
    source = exact_frames(sample, SOURCE_SYMBOL)
    rewritten = exact_frames(sample, REWRITTEN_SYMBOL)
    unknown = any(not frame.known for frame in sample.frames)
    wrong_dso = len(source_any) != len(source) or len(rewritten_any) != len(rewritten)
    if unknown:
        category = "ambiguous_unknown"
    elif wrong_dso:
        category = "ambiguous_symbol_dso"
    elif source and rewritten:
        category = "ambiguous_source_and_rewritten"
    elif source:
        category = "source_selected"
    elif rewritten:
        category = "rewritten_value_source"
    else:
        category = "neither"
    diagnostics = {
        "source_any_dso": len(source_any),
        "source_exact_binary": len(source),
        "rewritten_any_dso": len(rewritten_any),
        "rewritten_exact_binary": len(rewritten),
    }
    return category, diagnostics


def parse_symbol_offset(symbol: str) -> int | None:
    match = OFFSET_RE.search(symbol)
    return int(match.group(1), 16) if match else None


def child_paths(case_dir: Path) -> tuple[Path, Path, Path]:
    inline = case_dir / "offline-decode-v1/perf-script-inline.txt"
    physical = case_dir / "offline-decode-v2/perf-script-physical.txt"
    prior = case_dir / "offline-decode-v2/boundary-attribution-v1.json"
    return inline, physical, prior


def nm_symbols() -> dict[str, dict[str, int]]:
    command = ["nm", "-anS", "--demangle", "--defined-only", str(BINARY)]
    completed = subprocess.run(command, check=True, capture_output=True, text=True)
    found: dict[str, dict[str, int]] = {}
    for line in completed.stdout.splitlines():
        fields = line.split(maxsplit=3)
        if len(fields) != 4:
            continue
        address, size, _kind, name = fields
        if name not in TARGETS.values():
            continue
        found[name] = {"elf_start": int(address, 16), "size": int(size, 16)}
    require(len(found) == len(TARGETS), f"nm did not resolve all target symbols: {sorted(found)}")
    return found


def disassembly(symbol_meta: dict[str, dict[str, int]]) -> tuple[dict[str, dict[str, Any]], dict[str, Any]]:
    result: dict[str, dict[str, Any]] = {}
    commands: dict[str, Any] = {}
    for role, symbol in TARGETS.items():
        meta = symbol_meta[symbol]
        start = meta["elf_start"]
        stop = start + meta["size"]
        command = [
            "objdump",
            "-drwC",
            "--no-show-raw-insn",
            f"--start-address=0x{start:x}",
            f"--stop-address=0x{stop:x}",
            str(BINARY),
        ]
        completed = subprocess.run(command, check=True, capture_output=True, text=True)
        instructions: dict[int, str] = {}
        for line in completed.stdout.splitlines():
            match = DISASM_RE.fullmatch(line)
            if match:
                instructions[int(match.group(1), 16)] = match.group(2).strip()
        require(instructions, f"objdump returned no instructions for {symbol}")
        result[role] = {
            "symbol": symbol,
            "elf_start": start,
            "size": meta["size"],
            "instruction_count": len(instructions),
            "instructions": {f"0x{address - start:x}": text for address, text in instructions.items()},
        }
        commands[role] = command
    return result, commands


def addr2line(elf_address: int) -> list[str]:
    command = ["addr2line", "-e", str(BINARY), "-f", "-C", "-i", f"0x{elf_address:x}"]
    completed = subprocess.run(command, check=True, capture_output=True, text=True)
    return [line for line in completed.stdout.splitlines() if line]


def new_target_groups() -> dict[str, dict[str, int]]:
    return {role: empty_stats() for role in TARGETS}


def analyze_child(index: int, case_dir: Path, symbol_meta: dict[str, dict[str, int]]) -> dict[str, Any]:
    inline_path, physical_path, prior_path = child_paths(case_dir)
    with inline_path.open("r", encoding="utf-8", errors="strict") as stream:
        inline_samples, inline_diag = frozen.parse_samples(stream, EVENT)
    with physical_path.open("r", encoding="utf-8", errors="strict") as stream:
        physical_samples, physical_diag = frozen.parse_samples(stream, EVENT)
    prior = json.loads(prior_path.read_text(encoding="utf-8"))
    pairs, unmatched_inline, unmatched_physical, duplicate_keys = frozen.pair_samples(inline_samples, physical_samples)
    require(not unmatched_inline and not unmatched_physical, f"child {index}: paired streams differ")
    require(duplicate_keys == 0, f"child {index}: duplicate sample keys {duplicate_keys}")
    require(len(pairs) == len(inline_samples) == len(physical_samples), f"child {index}: pair count mismatch")

    selected = empty_stats()
    qualified = empty_stats()
    boundaries = {key: empty_stats() for key in ("source_selected", "rewritten_value_source", "neither", "ambiguous_source_and_rewritten", "ambiguous_unknown", "ambiguous_symbol_dso")}
    leaves = collections.defaultdict(empty_stats)
    target_self = new_target_groups()
    target_presence = new_target_groups()
    target_offsets: dict[str, dict[int | None, dict[str, int]]] = {role: collections.defaultdict(empty_stats) for role in TARGETS}
    target_runtime_ips: dict[str, dict[int | None, set[int]]] = {role: collections.defaultdict(set) for role in TARGETS}
    target_source_locations: dict[str, collections.Counter[str]] = {role: collections.Counter() for role in TARGETS}
    classification = collections.Counter()
    dso_mismatches = 0
    source_selected_pairs = 0
    source_selected_period = 0
    target_stack_depths: dict[str, collections.Counter[int]] = {role: collections.Counter() for role in TARGETS}
    source_file_evidence = collections.Counter()

    for inline_sample, physical_sample in pairs:
        cls, _reason, _evidence = frozen.classify_pair(
            inline_sample,
            physical_sample,
            source_suffix=frozen.DEFAULT_SOURCE_SUFFIX,
            stage_range=frozen.DEFAULT_STAGE_RANGE,
            timed_range=frozen.DEFAULT_TIMED_RANGE,
            physical_prefix=frozen.DEFAULT_PHYSICAL_TIMED_PREFIX,
            binary_dso=str(BINARY),
        )
        classification[cls] += 1
        if cls != "qualified":
            continue
        add_stats(qualified, inline_sample)
        category, diagnostics = classify_boundary(physical_sample)
        add_stats(boundaries[category], inline_sample)
        if category != "source_selected":
            continue
        source_selected_pairs += 1
        source_selected_period += inline_sample.period
        add_stats(selected, inline_sample)
        leaf = physical_sample.frames[0]
        leaf_symbol = exact_symbol(leaf) if leaf.known else "[unknown]"
        leaves[leaf_symbol]["samples"] += 1
        leaves[leaf_symbol]["period"] += inline_sample.period
        if diagnostics["source_any_dso"] != diagnostics["source_exact_binary"] or diagnostics["rewritten_any_dso"] != diagnostics["rewritten_exact_binary"]:
            dso_mismatches += 1
        if any(not frame.known for frame in physical_sample.frames):
            source_file_evidence["unknown_frame"] += 1
        else:
            source_file_evidence["known_frame_chain"] += 1
        for frame in physical_sample.frames:
            for role, symbol in TARGETS.items():
                if exact_symbol(frame) != symbol or frame.dso != str(BINARY):
                    continue
                target_stack_depths[role][frame.index] += 1
                if frame.index == 0:
                    add_stats(target_self[role], inline_sample)
                    offset = parse_symbol_offset(frame.symbol)
                    add_stats(target_offsets[role][offset], inline_sample)
                    target_runtime_ips[role][offset].add(frame.ip)
                    location = f"{frame.source_file}:{frame.source_line}" if frame.source_file and frame.source_line else "missing"
                    target_source_locations[role][location] += 1
                add_stats(target_presence[role], inline_sample)

    expected = prior["boundary_partition"]["source_selected"]
    require(selected == {"samples": expected["samples"], "period": expected["period"]}, f"child {index}: source-selected drift")
    require(qualified == {"samples": prior["selected_qualified"]["samples"], "period": prior["selected_qualified"]["period"]}, f"child {index}: qualified drift")
    require(sum(value["samples"] for value in boundaries.values()) == qualified["samples"], f"child {index}: boundary sample conservation")
    require(sum(value["period"] for value in boundaries.values()) == qualified["period"], f"child {index}: boundary period conservation")
    require(sum(value["samples"] for value in leaves.values()) == selected["samples"], f"child {index}: leaf sample conservation")
    require(sum(value["period"] for value in leaves.values()) == selected["period"], f"child {index}: leaf period conservation")
    return {
        "index": index,
        "case_directory": str(case_dir),
        "inputs": {
            "inline": descriptor(inline_path),
            "physical": descriptor(physical_path),
            "prior_boundary": descriptor(prior_path),
        },
        "qualified": qualified,
        "source_selected": selected,
        "boundary_partition": {key: value for key, value in sorted(boundaries.items())},
        "exclusive_physical_leaf": {key: value for key, value in sorted(leaves.items(), key=lambda item: (-item[1]["period"], item[0]))},
        "target_self": target_self,
        "target_presence": target_presence,
        "target_offsets": {
            role: {
                "offsets": {
                    ("missing" if offset is None else f"0x{offset:x}"): {
                        **value,
                        "runtime_ip_count": len(target_runtime_ips[role][offset]),
                        "runtime_ips": [f"0x{ip:x}" for ip in sorted(target_runtime_ips[role][offset])[:8]],
                    }
                    for offset, value in sorted(target_offsets[role].items(), key=lambda item: (-item[1]["period"], -1 if item[0] is None else item[0]))
                },
                "stack_depth_counts": {str(depth): count for depth, count in sorted(target_stack_depths[role].items())},
                "source_location_observed_only": target_source_locations[role].most_common(12),
            }
            for role in TARGETS
        },
        "source_evidence": {
            "selected_physical_pairs": source_selected_pairs,
            "selected_physical_period": source_selected_period,
            "dso_mismatch_pairs": dso_mismatches,
            "frame_chain_status": dict(source_file_evidence),
            "source_file_policy": "physical source_file/source_line are retained as observed evidence only; they are not used for owner qualification or coverage claims",
        },
        "qualification": {
            "classification_counts": dict(sorted(classification.items())),
            "inline_parser": inline_diag,
            "physical_parser": physical_diag,
            "duplicate_sample_keys": duplicate_keys,
        },
    }


def normalize_stats(value: dict[str, int], denominator: int) -> dict[str, Any]:
    return stats(value, denominator)


def aggregate(results: list[dict[str, Any]], symbol_meta: dict[str, dict[str, int]], disasm: dict[str, dict[str, Any]], disasm_commands: dict[str, Any]) -> dict[str, Any]:
    total_qualified = empty_stats()
    total_selected = empty_stats()
    boundaries = collections.defaultdict(empty_stats)
    leaves = collections.defaultdict(empty_stats)
    target_self = new_target_groups()
    target_presence = new_target_groups()
    target_offsets: dict[str, dict[str, dict[str, int]]] = {role: collections.defaultdict(empty_stats) for role in TARGETS}
    offset_ips: dict[str, dict[str, set[str]]] = {role: collections.defaultdict(set) for role in TARGETS}
    for result in results:
        for key in ("samples", "period"):
            total_qualified[key] += result["qualified"][key]
            total_selected[key] += result["source_selected"][key]
        for key, value in result["boundary_partition"].items():
            boundaries[key]["samples"] += value["samples"]
            boundaries[key]["period"] += value["period"]
        for key, value in result["exclusive_physical_leaf"].items():
            leaves[key]["samples"] += value["samples"]
            leaves[key]["period"] += value["period"]
        for role in TARGETS:
            for key in ("samples", "period"):
                target_self[role][key] += result["target_self"][role][key]
                target_presence[role][key] += result["target_presence"][role][key]
            for offset, value in result["target_offsets"][role]["offsets"].items():
                target_offsets[role][offset]["samples"] += value["samples"]
                target_offsets[role][offset]["period"] += value["period"]
                offset_ips[role][offset].update(value["runtime_ips"])
    require(total_qualified == EXPECTED_QUALIFIED, f"aggregate qualified mismatch: {total_qualified}")
    require(total_selected == EXPECTED_SOURCE_SELECTED, f"aggregate source-selected mismatch: {total_selected}")
    require(sum(value["samples"] for value in boundaries.values()) == total_qualified["samples"], "aggregate boundary sample conservation")
    require(sum(value["period"] for value in boundaries.values()) == total_qualified["period"], "aggregate boundary period conservation")
    require(sum(value["samples"] for value in leaves.values()) == total_selected["samples"], "aggregate leaf sample conservation")
    require(sum(value["period"] for value in leaves.values()) == total_selected["period"], "aggregate leaf period conservation")
    target_data: dict[str, Any] = {}
    for role, symbol in TARGETS.items():
        start = symbol_meta[symbol]["elf_start"]
        ranked_offsets = sorted(
            target_offsets[role].items(),
            key=lambda item: (-item[1]["period"], item[0]),
        )
        target_data[role] = {
            "symbol": symbol,
            "elf": symbol_meta[symbol],
            "exclusive_self": normalize_stats(target_self[role], total_selected["period"]),
            "inclusive_ancestry_presence": normalize_stats(target_presence[role], total_selected["period"]),
            "offsets": [
                {
                    "offset": offset,
                    "elf_address": (f"0x{start + int(offset, 16):x}" if offset != "missing" else None),
                    **normalize_stats(value, total_selected["period"]),
                    "runtime_ip_examples": sorted(offset_ips[role][offset])[:12],
                    "instruction": disasm[role]["instructions"].get(offset),
                    "instruction_boundary": offset == "missing" or offset in disasm[role]["instructions"],
                    "addr2line": (
                        addr2line(start + int(offset, 16))
                        if offset != "missing" and rank < MAX_ADDR2LINE_OFFSETS
                        else []
                    ),
                    "addr2line_bound": rank < MAX_ADDR2LINE_OFFSETS,
                }
                for rank, (offset, value) in enumerate(ranked_offsets)
            ],
        }
        for row in target_data[role]["offsets"]:
            require(row["instruction_boundary"], f"{role} offset {row['offset']} is not an objdump instruction boundary")
    return {
        "schema": SCHEMA,
        "status": "passed",
        "generated_utc": now(),
        "scope": {
            "source_category": "exact physical Snapshot::from_source_selected frame, binary DSO only, after frozen paired qualification",
            "exclusive_leaf": "deepest physical frame only; one sampled period charged once; callee costs excluded",
            "inclusive_presence": "diagnostic ancestry presence; target rows overlap and must not be summed",
            "source_file_policy": "source locations are observed evidence only; no source-file coverage or owner claim is made",
            "no_workload_or_build": True,
        },
        "binary": {
            **descriptor(BINARY),
            "expected_sha256": BINARY_SHA256,
            "gnu_build_id": BINARY_BUILD_ID,
        },
        "frozen_analyzer": {
            **descriptor(FROZEN_ANALYZER),
            "expected_sha256": FROZEN_ANALYZER_SHA256,
        },
        "denominators": {
            "qualified_wave": normalize_stats(total_qualified, total_qualified["period"]),
            "source_selected": normalize_stats(total_selected, total_selected["period"]),
        },
        "boundary_partition": {
            key: normalize_stats(value, total_qualified["period"])
            for key, value in sorted(boundaries.items())
        },
        "exclusive_physical_leaf_partition": {
            key: normalize_stats(value, total_selected["period"])
            for key, value in sorted(leaves.items(), key=lambda item: (-item[1]["period"], item[0]))
        },
        "targets": target_data,
        "disassembly": {
            role: {key: value for key, value in data.items() if key != "instructions"}
            for role, data in disasm.items()
        },
        "commands": {
            "nm": ["nm", "-anS", "--demangle", "--defined-only", str(BINARY)],
            "objdump": disasm_commands,
            "addr2line": ["addr2line", "-e", str(BINARY), "-f", "-C", "-i", "<elf-address>"],
        },
        "children": sorted(results, key=lambda item: item["index"]),
        "interpretation": {
            "namespace_semantic_scan": "process_event and resolve_event are quick-xml namespace-reader/resolver machinery; their self samples identify namespace transport/resolution work but do not prove removable semantics",
            "codegen_or_payload_movement": "instruction rows identify exact sampled machine instructions such as aggregate copies or branches; a code-generation hypothesis still needs source-preserving differential controls and ordinary-build measurement",
            "candidate_boundary": "The retained evidence supports investigating direct event transport in worksheet_xml_and_parse_source while preserving namespace push/pop, resolver timing, error order, limits, MCE handling, and borrowed event ownership. It does not establish benefit or authorize a source change.",
        },
    }


def write_new(path: Path, value: dict[str, Any]) -> None:
    require(not path.exists() and not path.is_symlink(), f"refusing to overwrite {path}")
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("x", encoding="utf-8") as stream:
        json.dump(value, stream, indent=2, sort_keys=True)
        stream.write("\n")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--wave-root", type=Path, default=DEFAULT_WAVE_ROOT)
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    try:
        require(sha256_file(FROZEN_ANALYZER) == FROZEN_ANALYZER_SHA256, "frozen analyzer hash changed")
        require(sha256_file(BINARY) == BINARY_SHA256, "exact binary hash changed")
        wave_root = args.wave_root.expanduser().absolute()
        case_dirs = sorted(wave_root.glob("[0-9][0-9]-*"))
        require(len(case_dirs) == 6, f"expected six wave children, found {len(case_dirs)}")
        symbol_meta = nm_symbols()
        disasm, disasm_commands = disassembly(symbol_meta)
        with concurrent.futures.ThreadPoolExecutor(max_workers=6) as pool:
            futures = [pool.submit(analyze_child, index, case_dir, symbol_meta) for index, case_dir in enumerate(case_dirs)]
            results = [future.result() for future in futures]
        receipt = aggregate(results, symbol_meta, disasm, disasm_commands)
        write_new(args.output.expanduser().absolute(), receipt)
        print(json.dumps({"schema": SCHEMA, "status": "passed", "output": str(args.output.expanduser().absolute())}, sort_keys=True))
        return 0
    except (AttributionError, OSError, UnicodeError, ValueError, json.JSONDecodeError, subprocess.SubprocessError) as exc:
        print(json.dumps({"schema": SCHEMA, "status": "failed", "error": str(exc)}, sort_keys=True), file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
