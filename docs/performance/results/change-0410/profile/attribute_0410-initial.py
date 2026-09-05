#!/usr/bin/env python3
"""Reproduce the 0410 XLSX selected-cell no-inline perf attribution.

This is a read-only postprocessor.  It consumes a completed ``perf script``
text export and checks the capture metadata/source revision before writing a
JSON summary.  A sample's perf period is retained as a weighted event-period
count; it is not wall time and it is not an operation-only counter.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import re
import subprocess
import sys
from collections import Counter
from pathlib import Path
from typing import Callable, Iterable

VERSION = "0410-attribution-v1"
DEFAULT_ROOT = Path("/home/zhuhe/code/litchi")
DEFAULT_CAPTURE = Path("/tmp/litchi-goal-0410-capture")
DEFAULT_ATTRIBUTION = Path("/tmp/litchi-goal-0410-attribution")
DEFAULT_SCRIPT = DEFAULT_ATTRIBUTION / "perf-script-no-inline.stdout"
DEFAULT_REPORT = DEFAULT_ATTRIBUTION / "perf-all-self.stdout"

# This is the exact demangled facade instantiation emitted by the captured
# release binary.  It is deliberately more specific than a bare substring
# such as ``SelectedWorksheet``.
TIMED_ANCESTOR_LITERAL = "<litchi::sheet::selection::SelectedWorksheet>::cell::<&str>"
SELECTED_SCAN_LITERAL = "litchi_xlsx::raw::worksheet::selected::scan_stream"
SOURCE_STREAM_LITERAL = "<litchi_xlsx::workbook::source::SourceWorksheet>::stream_cell"
SOURCE_CELL_LITERAL = "<litchi_xlsx::workbook::source::SourceWorksheet>::cell::<litchi_sheet::At>"
CHILD_LITERAL = "litchi_perf_baseline::filesystem::run_child_arguments"
X14AC_LITERAL = "litchi_xlsx::raw::worksheet::x14ac::capture_stream_with_active"

HEADER_RE = re.compile(
    r"^\S+\s+\d+\s+[^:]+:\s+(?P<period>\d+)\s+cycles(?::\S*)?\s*$"
)
FRAME_RE = re.compile(
    r"^\s*[0-9a-fA-F]+\s+(?P<symbol>.+?)\+0x[0-9a-fA-F]+\s+\([^)]*\)\s*$"
)
FRAME_RE_NO_OFFSET = re.compile(
    r"^\s*[0-9a-fA-F]+\s+(?P<symbol>.+?)\s+\([^)]*\)\s*$"
)


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as source:
        for chunk in iter(lambda: source.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def file_binding(path: Path, root: Path | None = None) -> dict[str, object]:
    resolved = path.resolve()
    item: dict[str, object] = {
        "path": str(resolved),
        "bytes": resolved.stat().st_size,
        "sha256": sha256_file(resolved),
    }
    if root is not None:
        try:
            item["relative_to_repo"] = str(resolved.relative_to(root.resolve()))
        except ValueError:
            pass
    return item


def git_blob(root: Path, revision: str, relative: str) -> bytes:
    return subprocess.check_output(
        ["git", "-C", str(root), "show", f"{revision}:{relative}"],
        stderr=subprocess.STDOUT,
    )


def git_source_binding(root: Path, revision: str) -> dict[str, object]:
    # These four files establish the measured facade, source-backed route,
    # selected stream scanner and child timing boundary.  Hash Git blobs from
    # the captured revision rather than the possibly dirty working tree.
    specs = [
        (
            "crates/litchi/src/sheet/selection.rs",
            [
                ("selected_facade_cell", re.compile(r"pub fn cell<'a>\(&self")),
            ],
        ),
        (
            "crates/litchi-xlsx/src/workbook/source.rs",
            [
                ("source_cell_route", re.compile(r"pub fn cell<'a>\(&self")),
                ("source_stream_cell", re.compile(r"fn stream_cell\(&self")),
                ("selected_scan_call", re.compile(r"raw::selected_worksheet::scan")),
            ],
        ),
        (
            "crates/litchi-xlsx/src/raw/worksheet/selected.rs",
            [
                ("selected_scan_stream", re.compile(r"fn scan_stream\(")),
                ("x14ac_capture_call", re.compile(r"capture_stream_with_active\(")),
            ],
        ),
        (
            "tools/perf-baseline/src/filesystem.rs",
            [
                ("operation_timer_start", re.compile(r"let started = Instant::now\(\);")),
                ("selected_operation_dispatch", re.compile(r"run_xlsx_operation\(operation, &source")),
                ("selected_cell_call", re.compile(r"let cell = sheet\.cell\(address\)")),
                ("post_timer_elapsed", re.compile(r"started\.elapsed\(\)\.as_nanos")),
                ("post_timer_verification", re.compile(r"verify_xlsx_operation\(")),
            ],
        ),
    ]
    files: list[dict[str, object]] = []
    for relative, checks in specs:
        blob = git_blob(root, revision, relative)
        lines = blob.decode("utf-8").splitlines()
        found: dict[str, list[int]] = {}
        for name, pattern in checks:
            found[name] = [n for n, line in enumerate(lines, 1) if pattern.search(line)]
        files.append(
            {
                "path": relative,
                "bytes": len(blob),
                "sha256": hashlib.sha256(blob).hexdigest(),
                "checks": found,
            }
        )
    return {
        "revision": revision,
        "files": files,
        "interpretation": {
            "timer": "filesystem.rs starts Instant before run_xlsx_operation and computes elapsed before post-operation verification",
            "selected_route": "SelectedWorksheet::cell -> SourceWorksheet::cell -> stream_cell -> selected::scan_stream -> x14ac capture",
        },
    }


def symbol_from_frame(line: str) -> str:
    match = FRAME_RE.match(line)
    if match:
        return match.group("symbol")
    match = FRAME_RE_NO_OFFSET.match(line)
    if match:
        return match.group("symbol")
    # Keep an unrecognised frame visible in the summary.  It can still be
    # matched as a marker if perf changes its formatting.
    return line.strip()


def is_function(symbol: str, name: str) -> bool:
    """Match a complete terminal Rust function name, avoiding name_part."""
    return symbol == name or symbol.endswith(f"::{name}")


def parse_samples(path: Path) -> tuple[list[tuple[int, list[str]]], dict[str, int]]:
    samples: list[tuple[int, list[str]]] = []
    malformed_headers = 0
    nonempty_before_header = 0
    header: str | None = None
    period = 0
    frame_lines: list[str] = []

    def flush() -> None:
        nonlocal header, period, frame_lines
        if header is not None:
            samples.append((period, [symbol_from_frame(line) for line in frame_lines]))
        header = None
        period = 0
        frame_lines = []

    with path.open("r", encoding="utf-8", errors="replace") as source:
        for raw in source:
            line = raw.rstrip("\n")
            match = HEADER_RE.match(line)
            if match:
                flush()
                header = line
                period = int(match.group("period"))
            elif line.strip() == "":
                continue
            elif header is None:
                nonempty_before_header += 1
                # perf script can emit metadata in some modes.  This capture
                # has none; retain a counter instead of silently accepting it.
            else:
                frame_lines.append(line)
    flush()
    # Header shape is strict enough that a line containing cycles but not
    # parsed as a header should be visible in the output rather than dropped.
    with path.open("r", encoding="utf-8", errors="replace") as source:
        for raw in source:
            if "cycles" in raw and not HEADER_RE.match(raw.rstrip("\n")):
                malformed_headers += 1
    return samples, {
        "sample_blocks": len(samples),
        "malformed_cycle_headers": malformed_headers,
        "nonempty_lines_before_header": nonempty_before_header,
    }


def has(symbols: Iterable[str], literal: str) -> bool:
    return any(literal in symbol for symbol in symbols)


def first_index(symbols: list[str], predicate: Callable[[str], bool]) -> int | None:
    for index, symbol in enumerate(symbols):
        if predicate(symbol):
            return index
    return None


def weighted_metric(
    periods: int, blocks: int, denominator: int, whole: int
) -> dict[str, object]:
    return {
        "weighted_event_period": periods,
        "raw_stack_blocks": blocks,
        "share_of_timed_percent": (periods / denominator * 100.0) if denominator else None,
        "share_of_whole_process_percent": (periods / whole * 100.0) if whole else None,
    }


def aggregate(
    samples: list[tuple[int, list[str]]],
    predicate: Callable[[list[str]], bool],
    timed_period: int,
    whole_period: int,
) -> dict[str, object]:
    period = 0
    blocks = 0
    for sample_period, symbols in samples:
        if predicate(symbols):
            period += sample_period
            blocks += 1
    return weighted_metric(period, blocks, timed_period, whole_period)


def output_commands(capture: Path, attribution: Path) -> dict[str, object]:
    commands_path = capture / "commands.json"
    data = json.loads(commands_path.read_text())
    recorded = {
        item["label"]: item
        for item in data
        if item.get("label") in {"selected-perf-record", "perf-all-self", "perf-script-no-inline"}
    }
    return {
        "capture_commands_json": file_binding(commands_path),
        "recorded_capture_argv": {
            label: item.get("argv") for label, item in recorded.items()
        },
        "supplementary_no_inline_argv": [
            [
                "perf",
                "report",
                "--stdio",
                "--no-inline",
                "--no-children",
                "--call-graph=none",
                "--percent-limit=0",
                "-i",
                str(capture / "perf.data"),
            ],
            [
                "perf",
                "script",
                "--no-inline",
                "-i",
                str(capture / "perf.data"),
            ],
        ],
        "supplementary_command_provenance": "argv, exit status and command timing are retained in commands.json; postprocessing duration is not API latency",
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--script", type=Path, default=DEFAULT_SCRIPT)
    parser.add_argument("--report", type=Path, default=DEFAULT_REPORT)
    parser.add_argument("--capture", type=Path, default=DEFAULT_CAPTURE)
    parser.add_argument("--repo", type=Path, default=DEFAULT_ROOT)
    parser.add_argument(
        "--output",
        type=Path,
        default=DEFAULT_ATTRIBUTION / "attribution-summary.json",
    )
    args = parser.parse_args()
    script_path = args.script.resolve()
    report_path = args.report.resolve()
    capture = args.capture.resolve()
    repo = args.repo.resolve()
    output_path = args.output.resolve()
    for required in [script_path, report_path, capture / "environment.json", capture / "commands.json", capture / "perf.data"]:
        if not required.is_file():
            raise SystemExit(f"missing required input: {required}")

    environment_path = capture / "environment.json"
    environment = json.loads(environment_path.read_text())
    source_before = environment["source_before"]
    revision = source_before["revision"]
    script_hash = sha256_file(script_path)
    samples, parse_shape = parse_samples(script_path)
    whole_period = sum(period for period, _ in samples)
    timed_samples = [
        (period, symbols)
        for period, symbols in samples
        if TIMED_ANCESTOR_LITERAL in symbols
    ]
    timed_period = sum(period for period, _ in timed_samples)
    if not samples:
        raise SystemExit("no perf sample blocks parsed")
    if not timed_samples:
        raise SystemExit(f"timed ancestor was not found: {TIMED_ANCESTOR_LITERAL}")

    def timed_predicate(predicate: Callable[[list[str]], bool]) -> Callable[[list[str]], bool]:
        return lambda symbols: TIMED_ANCESTOR_LITERAL in symbols and predicate(symbols)

    # Ordered stack check is leaf -> caller (perf script's printed order).
    callpath_predicates: list[tuple[str, Callable[[str], bool]]] = [
        ("x14ac_capture", lambda s: X14AC_LITERAL in s),
        ("selected_scan_stream", lambda s: SELECTED_SCAN_LITERAL in s),
        ("source_stream_cell", lambda s: SOURCE_STREAM_LITERAL in s),
        ("source_cell", lambda s: SOURCE_CELL_LITERAL in s),
        ("selected_facade_cell", lambda s: s == TIMED_ANCESTOR_LITERAL),
        ("filesystem_child", lambda s: CHILD_LITERAL in s),
    ]
    path_counts: dict[str, object] = {}
    for name, predicate in callpath_predicates:
        path_counts[name] = aggregate(
            samples,
            lambda symbols, predicate=predicate: TIMED_ANCESTOR_LITERAL in symbols
            and any(predicate(symbol) for symbol in symbols),
            timed_period,
            whole_period,
        )

    ordered_callpath = {
        "leaf_to_caller": [
            "MCE/parser leaf (varies)",
            X14AC_LITERAL,
            SELECTED_SCAN_LITERAL,
            SOURCE_STREAM_LITERAL,
            SOURCE_CELL_LITERAL,
            TIMED_ANCESTOR_LITERAL,
            CHILD_LITERAL,
        ],
        "timed_blocks_with_each_marker": path_counts,
        "ordered_chain_blocks": 0,
        "ordered_chain_weighted_event_period": 0,
    }
    ordered_period = 0
    ordered_blocks = 0
    for period, symbols in timed_samples:
        positions = []
        for _, predicate in callpath_predicates:
            position = first_index(symbols, predicate)
            positions.append(position)
        if all(position is not None for position in positions) and positions == sorted(positions):
            ordered_period += period
            ordered_blocks += 1
    ordered_callpath["ordered_chain_blocks"] = ordered_blocks
    ordered_callpath["ordered_chain_weighted_event_period"] = ordered_period
    ordered_callpath["ordered_chain_share_of_timed_percent"] = ordered_period / timed_period * 100.0

    # Exact leaf accounting is based on frame zero, while presence accounting
    # is inclusive stack ancestry.  This avoids conflating clone_bounded_name
    # with the longer clone_bounded_name_part symbol.
    targets = {
        "processor_parse_element": lambda s: is_function(s, "parse_element"),
        "clone_bounded_name_part": lambda s: is_function(s, "clone_bounded_name_part"),
        "clone_bounded_name": lambda s: is_function(s, "clone_bounded_name"),
        "clone_bounded_bytes": lambda s: is_function(s, "clone_bounded_bytes"),
        "clone_bounded_text": lambda s: is_function(s, "clone_bounded_text"),
        "clone_bounded_string": lambda s: is_function(s, "clone_bounded_string"),
        "expand": lambda s: is_function(s, "expand"),
        "semantic_attrs": lambda s: is_function(s, "semantic_attrs"),
        "clone_raw_attributes": lambda s: is_function(s, "clone_raw_attributes"),
        "validate_duplicate_attributes": lambda s: is_function(s, "validate_duplicate_attributes"),
        "namespaces_with_local": lambda s: is_function(s, "with_local"),
        "raw_start_commit": lambda s: is_function(s, "raw_start_commit"),
        "close_start": lambda s: is_function(s, "close_start"),
        "semantic_event_from_element": lambda s: is_function(s, "from_element"),
    }
    target_metrics: dict[str, object] = {}
    for name, predicate in targets.items():
        presence = aggregate(samples, lambda symbols, predicate=predicate: any(predicate(s) for s in symbols) and TIMED_ANCESTOR_LITERAL in symbols, timed_period, whole_period)
        leaf = aggregate(samples, lambda symbols, predicate=predicate: bool(symbols) and predicate(symbols[0]) and TIMED_ANCESTOR_LITERAL in symbols, timed_period, whole_period)
        target_metrics[name] = {"presence_in_timed_stack": presence, "exact_leaf": leaf}

    # Immediate caller of each exact leaf is the next frame in perf's
    # leaf-to-root order.  This is the Name-vs-lexical-copy discriminator.
    caller_breakdowns: dict[str, object] = {}
    for target_name, predicate in targets.items():
        callers: Counter[str] = Counter()
        for period, symbols in timed_samples:
            if symbols and predicate(symbols[0]):
                callers[symbols[1] if len(symbols) > 1 else "<none>"] += period
        base = sum(callers.values())
        caller_breakdowns[target_name] = [
            {
                "caller": caller,
                **weighted_metric(period, 0, timed_period, whole_period),
                "share_of_leaf_percent": period / base * 100.0 if base else None,
            }
            for caller, period in callers.most_common(12)
        ]
        for item in caller_breakdowns[target_name]:
            # raw blocks are needed for the caller breakdown too.
            item["raw_stack_blocks"] = sum(
                1
                for sample_period, symbols in timed_samples
                if symbols
                and predicate(symbols[0])
                and (symbols[1] if len(symbols) > 1 else "<none>") == item["caller"]
            )

    # The leaf table is intentionally limited but complete enough to expose
    # the dominant selected scan leaves without carrying all 80K unique names.
    leaves: Counter[str] = Counter()
    leaf_blocks: Counter[str] = Counter()
    for period, symbols in timed_samples:
        leaf = symbols[0] if symbols else "<no-frame>"
        leaves[leaf] += period
        leaf_blocks[leaf] += 1
    top_leaves = [
        {
            "symbol": symbol,
            **weighted_metric(period, leaf_blocks[symbol], timed_period, whole_period),
        }
        for symbol, period in leaves.most_common(40)
    ]

    # Selected scan is a useful narrower subset because a handful of timed
    # samples occur at the facade/caller while no scan frame is visible.
    selected_scan = aggregate(
        samples,
        timed_predicate(lambda symbols: SELECTED_SCAN_LITERAL in "\n".join(symbols)),
        timed_period,
        whole_period,
    )
    x14ac = aggregate(
        samples,
        timed_predicate(lambda symbols: X14AC_LITERAL in "\n".join(symbols)),
        timed_period,
        whole_period,
    )
    parse_presence = target_metrics["processor_parse_element"]["presence_in_timed_stack"]
    lexical_bytes = target_metrics["clone_bounded_bytes"]
    lexical_text = target_metrics["clone_bounded_text"]
    name_part = target_metrics["clone_bounded_name_part"]

    # Validate the same identity facts that make this attribution interpretable.
    normal = json.loads((capture / "selected-normal.json").read_text())
    allocator = json.loads((capture / "selected-allocator.json").read_text())
    perf_record = json.loads((capture / "selected-perf-record.json").read_text())
    expected_corpus_sha = "dfff7ec0c749d9e404101776f15a8fb690985af7f58efdfe659dbeaed7145036"
    identity_checks = {
        "capture_revision": revision,
        "worktree_status": source_before["status"],
        "clean_worktree": source_before["status"] == "",
        "normal_case": normal["results"][0]["case"],
        "normal_samples": len(normal["results"][0]["elapsed_ns"]["samples"]),
        "normal_warmup": normal["configuration"]["warmup_iterations_per_case"],
        "allocator_calls": allocator["results"][0]["operation_metrics"]["allocation"]["allocation_calls"],
        "allocator_bytes": allocator["results"][0]["operation_metrics"]["allocation"]["allocated_bytes"],
        "perf_record_case": perf_record["results"][0]["case"],
        "perf_record_samples": len(perf_record["results"][0]["elapsed_ns"]["samples"]),
        "corpus_archive_sha256": normal["results"][0]["corpus"]["archive_sha256"],
        "selected_cell_digest": normal["filesystem_evidence"][0]["samples"][0]["xlsx_selected_cell"]["digest"],
    }
    if identity_checks["normal_case"] != "xlsx_file_selected_cell":
        raise SystemExit("capture case identity mismatch")
    if identity_checks["corpus_archive_sha256"] != expected_corpus_sha:
        raise SystemExit("corpus archive identity mismatch")
    if identity_checks["perf_record_samples"] != 300:
        raise SystemExit("perf record sample identity mismatch")
    if identity_checks["normal_samples"] != 500 or identity_checks["normal_warmup"] != 20:
        raise SystemExit("normal sample identity mismatch")

    relevant_files = [
        environment_path,
        capture / "commands.json",
        capture / "selected-normal.json",
        capture / "selected-normal.catalog.json",
        capture / "selected-allocator.json",
        capture / "selected-allocator.catalog.json",
        capture / "selected-perf-record.json",
        capture / "selected-perf-record.catalog.json",
        capture / "perf.data",
        script_path,
        report_path,
        args.script.with_name("perf-script-no-inline.stderr").resolve(),
        args.report.with_name("perf-all-self.stderr").resolve(),
    ]
    input_bindings = []
    for path in relevant_files:
        if path.is_file():
            input_bindings.append(file_binding(path, repo))

    source_binding = git_source_binding(repo, revision)
    summary: dict[str, object] = {
        "schema": VERSION,
        "purpose": "Current-state descriptive CPU attribution for xlsx_file_selected_cell; no before/after or speedup claim",
        "inputs": {
            "perf_script": file_binding(script_path, repo),
            "perf_report": file_binding(report_path, repo),
            "capture_artifacts": input_bindings,
            "parser": file_binding(Path(__file__).resolve(), repo),
        },
        "capture_identity": identity_checks,
        "environment": environment,
        "source_binding": source_binding,
        "sample_parser": {
            **parse_shape,
            "whole_process_weighted_event_period": whole_period,
            "whole_process_raw_stack_blocks": len(samples),
            "event": "cycles:u",
            "weight_semantics": "perf script sample period; weighted event period, not wall time",
        },
        "timed_scope": {
            "case": "xlsx_file_selected_cell",
            "ancestor_literal": TIMED_ANCESTOR_LITERAL,
            "ancestor_match": "exact complete demangled symbol after offset removal",
            "weighted_event_period": timed_period,
            "raw_stack_blocks": len(timed_samples),
            "share_of_whole_process_percent": timed_period / whole_period * 100.0,
            "selected_scan": selected_scan,
            "x14ac_capture": x14ac,
            "scope_note": "perf samples are whole process and inherited children; this ancestor subset is a stack-attributed selector sample subset, not an independent timer or wall-phase total",
        },
        "callpath_verification": ordered_callpath,
        "target_metrics": target_metrics,
        "immediate_leaf_callers": caller_breakdowns,
        "top_timed_leafs": top_leaves,
        "lexical_vs_name_interpretation": {
            "expanded_name_clone": name_part,
            "expanded_name_clone_callers": caller_breakdowns["clone_bounded_name_part"],
            "lexical_bytes_clone": lexical_bytes,
            "lexical_text_clone": lexical_text,
            "lexical_text_resolution": "no standalone clone_bounded_text frame in no-inline output; this is unresolved/inlined, not evidence of zero cost",
            "shared_allocator_leafs": "malloc, memmove, and deallocation symbols are shared across parser, expanded Name, event, frame, and verification paths and are not uniquely assigned to lexical copies",
        },
        "limitations": [
            "Perf data includes process startup, per-child corpus reconstruction, archive verification, and post-operation semantic verification; those are excluded only by the selected-cell ancestor stack filter, not by a separate timestamp boundary in perf.",
            "The selected-cell operation scans the source worksheet stream to EOF; the selected ancestor subset therefore includes the selected stream parser and its MCE/x14ac callbacks.",
            "The source tree was dirty at capture because docs/GOAL.md was intentionally untracked; the binary and source binding are to the recorded revision, and the tracked files used for the route checks are hashed from that Git revision.",
            "A single current-state capture supports hotspot evidence only. It does not support an optimization speedup, regression, or ABBA claim.",
        ],
        "recommendation": "The stronger measured named leaf is expanded Name cloning (clone_bounded_name_part), mostly called by clone_bounded_name and expand. Lexical clone_bounded_bytes is only 0.038% of timed weighted periods as an exact leaf and 0.361% as stack presence; authorize only a narrow borrowing experiment after correctness and fresh profiling, not a broad lexical-copy bottleneck claim.",
        "commands": output_commands(capture, args.script.parent),
    }
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(json.dumps(summary, indent=2, sort_keys=True) + "\n")
    print(json.dumps({
        "output": str(output_path),
        "whole_process_period": whole_period,
        "whole_process_blocks": len(samples),
        "timed_period": timed_period,
        "timed_blocks": len(timed_samples),
        "parse_element_presence_period": parse_presence["weighted_event_period"],
        "parse_element_leaf_period": target_metrics["processor_parse_element"]["exact_leaf"]["weighted_event_period"],
        "name_part_leaf_period": name_part["exact_leaf"]["weighted_event_period"],
        "lexical_bytes_leaf_period": lexical_bytes["exact_leaf"]["weighted_event_period"],
    }, sort_keys=True))
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except subprocess.CalledProcessError as error:
        raise SystemExit(f"git source binding failed: {error}")
