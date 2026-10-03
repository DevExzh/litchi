#!/usr/bin/env python3
"""Annotate and analyze the separate current-source Callgrind captures.

The selected rows are exact canonical function/file rows.  They are inclusive
guest-instruction observations, may overlap, and never become native time or
speedup estimates.
"""

from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path
import re
import subprocess
from typing import Any

import verify


P = Path(__file__).resolve().parent
SELECTORS = {
    "runner": "tools/perf-baseline/src/producer_shape.rs:litchi_perf_baseline::producer_shape::run_xlsx_full_cell_scan [",
    "mce_process": "crates/litchi-ooxml-common/src/mce/codec.rs:litchi_ooxml_common::mce::codec::process_markup_compatibility [",
    "mce_start": "crates/litchi-ooxml-common/src/mce/codec.rs:litchi_ooxml_common::mce::codec::start [",
    "x14ac_capture": "crates/litchi-xlsx/src/raw/worksheet/x14ac.rs:litchi_xlsx::raw::worksheet::x14ac::capture_inner [",
}


def annotate(folder: Path) -> tuple[Path, Path]:
    raw = folder / "callgrind.out"
    canonical = folder / "inclusive-canonical.txt"
    stderr = folder / "annotate.stderr"
    command_receipt = folder / "annotate-command.json"
    if canonical.exists() or stderr.exists() or command_receipt.exists():
        raise FileExistsError(f"Callgrind annotation already exists; preserve prior attempt: {folder}")
    argv = ["/usr/bin/callgrind_annotate", "--inclusive=yes", "--threshold=100", "--auto=no", str(raw)]
    with canonical.open("x", encoding="utf-8") as out, stderr.open("x", encoding="utf-8") as err:
        code = subprocess.run(argv, cwd=P, env=verify.CHILD_ENV, stdout=out, stderr=err, check=False).returncode
    verify.write_once(command_receipt, {
        "schema": "litchi.xlsx-scan-callgrind-annotate-command-0892.v1",
        "argv": argv,
        "exit_code": code,
        "output": verify.descriptor(canonical),
        "stderr": verify.descriptor(stderr),
        "helper": verify.descriptor(Path(__file__)),
    })
    if code != 0:
        raise RuntimeError(f"callgrind_annotate failed for {folder}; outputs retained")
    return canonical, stderr


def parse_child(folder: Path) -> dict[str, Any]:
    raw = (folder / "callgrind.out").read_text(encoding="utf-8", errors="replace")
    annotation_path = folder / "inclusive-canonical.txt"
    stderr_path = folder / "stderr.txt"
    if not annotation_path.is_file():
        annotate(folder)
    annotation = annotation_path.read_text(encoding="utf-8", errors="replace")
    stderr = stderr_path.read_text(encoding="utf-8", errors="replace") if stderr_path.is_file() else ""
    totals = re.findall(r"^totals:\s+(\d+)", raw, re.MULTILINE)
    summaries = re.findall(r"^summary:\s+(\d+)", raw, re.MULTILINE)
    require_total = totals and summaries and int(totals[-1]) > 0 and int(totals[-1]) == int(summaries[0])
    if not require_total:
        raise ValueError(f"Callgrind instruction totals disagree: {folder.name}")
    total = int(totals[-1])
    pids = re.findall(r"^pid:\s+(\d+)", raw, re.MULTILINE)
    callgrind_pids = re.findall(r"^==(\d+)== Callgrind", stderr, re.MULTILINE)
    if not pids or not callgrind_pids or int(pids[0]) != int(callgrind_pids[0]):
        raise ValueError(f"raw profile is not the original parent process: {folder.name}")
    selected: dict[str, Any] = {}
    for name, selector in SELECTORS.items():
        matches = [line for line in annotation.splitlines() if selector in line]
        if len(matches) != 1:
            raise ValueError(f"canonical function row is missing or ambiguous: {folder.name}/{name}")
        match = re.match(r"^\s*([\d,]+)\s+", matches[0])
        if match is None:
            raise ValueError(f"canonical row has no instruction count: {folder.name}/{name}")
        value = int(match.group(1).replace(",", ""))
        selected[name] = {
            "inclusive_ir": value,
            "percent_of_runner": 100 * value / total,
            "selector": selector,
        }
    if selected["runner"]["inclusive_ir"] != total:
        raise ValueError(f"selected runner does not cover collected instructions: {folder.name}")
    return {
        "child": folder.name,
        "total_ir": total,
        "selected_inclusive_rows": selected,
        "parent_pid": int(pids[0]),
        "brk_warning": "brk segment overflow" in stderr,
        "callgrind_sha256": hashlib.sha256(raw.encode()).hexdigest(),
        "annotation_sha256": hashlib.sha256(annotation.encode()).hexdigest(),
        "annotation_command": verify.read(folder / "annotate-command.json") if (folder / "annotate-command.json").is_file() else None,
    }


def compare(rows: list[dict[str, Any]], baseline_path: Path) -> dict[str, Any]:
    baseline = verify.read(baseline_path)
    old_rows = baseline.get("rows")
    if not isinstance(old_rows, list) or len(old_rows) != len(rows):
        raise ValueError("0886 Callgrind analysis row count mismatch")
    comparisons = []
    for current, old in zip(rows, old_rows):
        old_selected = old.get("selected_inclusive_rows", {})
        values: dict[str, Any] = {}
        for name, item in current["selected_inclusive_rows"].items():
            previous = old_selected.get(name)
            if not isinstance(previous, dict):
                continue
            values[name] = {
                "current_inclusive_ir": item["inclusive_ir"],
                "baseline_0886_inclusive_ir": previous["inclusive_ir"],
                "current_percent_of_runner": item["percent_of_runner"],
                "baseline_0886_percent_of_runner": previous["percent_of_runner"],
                "percent_delta": item["percent_of_runner"] - previous["percent_of_runner"],
            }
        comparisons.append({"child": current["child"], "current_total_ir": current["total_ir"], "baseline_0886_total_ir": old.get("total_ir"), "rows": values})
    return {
        "baseline": {"path": str(baseline_path.resolve()), "sha256": verify.sha256(baseline_path)},
        "rows": comparisons,
        "overlap": ["mce_start is nested in mce_process; selected inclusive rows must not be summed"],
        "interpretation": "Guest instruction attribution only; no native inclusive time, allocation, latency, or speedup claim.",
    }


def analyze(capture_root: Path, baseline_path: Path, output: Path) -> dict[str, Any]:
    if output.exists():
        raise FileExistsError(f"Callgrind analysis output already exists; preserve prior attempt: {output}")
    folders = sorted(path for path in capture_root.glob("callgrind-*") if path.is_dir())
    if len(folders) != verify.CALLGRIND_CHILDREN:
        raise ValueError(f"expected {verify.CALLGRIND_CHILDREN} Callgrind children, found {len(folders)}")
    rows = [parse_child(folder) for folder in folders]
    result = {
        "schema": "litchi.xlsx-scan-current-runner-ir-0892.v1",
        "status": "pass",
        "rows": rows,
        "comparison_to_0886": compare(rows, baseline_path),
        "scope": "Two separately captured 10-sample/zero-warmup runner-scoped guest instruction diagnostics. Inclusive rows overlap; never sum rows or combine with native cycles.",
        "performance_claim": False,
    }
    verify.write_once(output, result)
    return result


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--capture-root", type=Path, default=P / "callgrind")
    parser.add_argument("--baseline", type=Path, default=P.parent / "change-0886" / "callgrind-analysis.json")
    parser.add_argument("--output", type=Path, default=P / "callgrind-analysis.json")
    args = parser.parse_args()
    result = analyze(args.capture_root.resolve(), args.baseline.resolve(), args.output.resolve())
    print(json.dumps(result, indent=2, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
