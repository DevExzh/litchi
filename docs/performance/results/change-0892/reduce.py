#!/usr/bin/env python3
"""Reduce current native profiles to disjoint sampled-IP self attribution.

The reducer reconciles every sampled IP period and count against ``perf
report``.  It compares descriptive self-symbol shares with the retained 0886
profile and deliberately emits no native-inclusive attribution or speedup
claim.
"""

from __future__ import annotations

import argparse
import collections
import json
from pathlib import Path
import re
from typing import Any

import verify


P = Path(__file__).resolve().parent
LEAF = re.compile(r"^\d+/\d+\s+[\d.]+:\s+(\d+)\s+(\S+):\s+([0-9a-f]+)\s+(.*?)\s+\((.*)\)$")
HEADER = re.compile(r"^\d+/\d+\s+[\d.]+:\s+(\d+)\s+(\S+):")


def key(dso: str, symbol: str) -> tuple[str, str]:
    if dso == "[unknown]":
        return ("[unknown]", "[unknown]")
    if symbol == "[unknown]" or re.fullmatch(r"0x[0-9a-f]+", symbol):
        return (Path(dso).name, "[unknown]")
    return (Path(dso).name, symbol)


def group(dso: str, symbol: str) -> str:
    if dso == "[unknown]" or symbol == "[unknown]" or symbol.startswith("0x"):
        return "unresolved_sampled_ip"
    for prefix, name in (
        ("litchi_ooxml_common::mce::", "mce"),
        ("litchi_xlsx::raw::worksheet::x14ac::", "x14ac"),
        ("quick_xml::name::NamespaceResolver::", "xml_namespace_resolver"),
        ("quick_xml::reader::ns_reader::", "xml_namespace_reader"),
        ("quick_xml::", "other_quick_xml"),
        ("litchi_xlsx::raw::worksheet::", "other_worksheet_parser"),
        ("zlib_rs::", "zlib"),
    ):
        if symbol.startswith(prefix):
            return name
    return "other_resolved"


def parse_child(folder: Path) -> dict[str, Any]:
    leaves = collections.Counter()
    leaf_counts = collections.Counter()
    for raw_line in (folder / "leaf-samples.txt").read_text(encoding="utf-8", errors="replace").splitlines():
        line = raw_line.strip()
        match = LEAF.fullmatch(line)
        if match is None:
            raise ValueError(f"unparsed sampled IP in {folder.name}: {raw_line}")
        period, event, _ip, symbol, dso = match.groups()
        if event.strip() != verify.EVENT:
            raise ValueError(f"unexpected event in {folder.name}: {event!r}")
        item = key(dso.strip(), symbol.strip())
        leaves[item] += int(period)
        leaf_counts[item] += 1
    report = (folder / "self-symbols.txt").read_text(encoding="utf-8", errors="replace")
    if "# Total Lost Samples: 0" not in report:
        raise ValueError(f"lost samples require investigation: {folder.name}")
    reported = collections.Counter()
    reported_counts = collections.Counter()
    for raw_line in report.splitlines():
        if raw_line.lstrip().startswith("#") or "|" not in raw_line:
            continue
        fields = [value.strip() for value in raw_line.split("|")]
        if len(fields) < 5 or not fields[1].isdigit() or not fields[2].replace(",", "").isdigit():
            continue
        _overhead, count, period, dso, symbol = fields[:5]
        symbol = re.sub(r"^\[.\]\s+", "", symbol)
        item = key(dso, symbol)
        reported[item] += int(period.replace(",", ""))
        reported_counts[item] += int(count)
    if leaves != reported or leaf_counts != reported_counts:
        raise ValueError(f"sampled-IP and self-symbol views disagree: {folder.name}")
    total = sum(leaves.values())
    event_match = re.search(r"Event count \(approx\.\):\s*([0-9,]+)", report)
    if event_match is None or total != int(event_match.group(1).replace(",", "")):
        raise ValueError(f"event period total mismatch: {folder.name}")
    groups = collections.Counter()
    for (dso, symbol), period in leaves.items():
        groups[group(dso, symbol)] += period
    stack_period = 0
    empty_period = 0
    unknown_stack_period = 0
    stack_text = (folder / "call-stacks.txt").read_text(encoding="utf-8", errors="replace").strip()
    if stack_text:
        for block in stack_text.split("\n\n"):
            lines = block.splitlines()
            if not lines:
                continue
            match = HEADER.match(lines[0].strip())
            if match is None or match.group(2).strip() != verify.EVENT:
                raise ValueError(f"unparsed stack sample in {folder.name}: {lines[0]}")
            period = int(match.group(1))
            stack_period += period
            if len(lines) == 1:
                empty_period += period
            if any("[unknown]" in line for line in lines[1:]):
                unknown_stack_period += period
    if stack_period != total:
        raise ValueError(f"stack sample population differs: {folder.name}")
    return {
        "child": folder.name,
        "sample_count": sum(leaf_counts.values()),
        "total_period": total,
        "lost_samples": 0,
        "self_groups": {name: {"period": period, "percent": period / total * 100} for name, period in groups.items()},
        "empty_stack_percent": empty_period / total * 100,
        "any_unknown_stack_percent": unknown_stack_period / total * 100,
        "top_self_symbols": [
            {"dso": item[0], "symbol": item[1], "period": period, "samples": leaf_counts[item], "percent": period / total * 100}
            for item, period in leaves.most_common(25)
        ],
        "input_sha256": {
            name: verify.sha256(folder / name)
            for name in ("leaf-samples.txt", "self-symbols.txt", "call-stacks.txt")
        },
    }


def _group_ranges(rows: list[dict[str, Any]]) -> dict[str, dict[str, float]]:
    values: dict[str, list[float]] = collections.defaultdict(list)
    for row in rows:
        for name, value in row["self_groups"].items():
            values[name].append(float(value["percent"]))
    return {
        name: {"min_percent": min(items), "max_percent": max(items), "mean_percent": sum(items) / len(items)}
        for name, items in sorted(values.items())
    }


def _comparison(rows: list[dict[str, Any]], baseline_path: Path) -> dict[str, Any]:
    baseline = verify.read(baseline_path)
    baseline_rows = baseline.get("rows")
    if not isinstance(baseline_rows, list) or not baseline_rows:
        raise ValueError("0886 baseline self-analysis has no rows")
    current_ranges = _group_ranges(rows)
    baseline_ranges = _group_ranges(baseline_rows)
    groups = sorted(set(current_ranges) | set(baseline_ranges))
    table: dict[str, Any] = {}
    for name in groups:
        current = current_ranges.get(name, {"min_percent": 0.0, "max_percent": 0.0, "mean_percent": 0.0})
        old = baseline_ranges.get(name, {"min_percent": 0.0, "max_percent": 0.0, "mean_percent": 0.0})
        table[name] = {
            "current": current,
            "baseline_0886": old,
            "mean_delta_percentage_points": current["mean_percent"] - old["mean_percent"],
        }
    return {
        "baseline": {"path": str(baseline_path.resolve()), "sha256": verify.sha256(baseline_path)},
        "groups": table,
        "interpretation": "Descriptive disjoint sampled-IP self shares only; incomplete stacks prevent caller or native-inclusive attribution. No profiler speedup claim.",
    }


def reduce_profiles(capture_root: Path, baseline_path: Path, output: Path) -> dict[str, Any]:
    if output.exists():
        raise FileExistsError(f"reduction output already exists; preserve prior attempt: {output}")
    folders = sorted(path for path in capture_root.glob("child-*") if path.is_dir())
    if len(folders) != verify.NATIVE_CHILDREN:
        raise ValueError(f"expected {verify.NATIVE_CHILDREN} native children, found {len(folders)}")
    rows = [parse_child(folder) for folder in folders]
    result = {
        "schema": "litchi.xlsx-scan-current-self-analysis-0892.v1",
        "status": "pass",
        "rows": rows,
        "comparison_to_0886": _comparison(rows, baseline_path),
        "scope": "Exact period-weighted sampled-IP self attribution; categories are disjoint. No inclusive or caller cost inferred from incomplete stacks. Guest instructions remain separate.",
        "performance_claim": False,
    }
    verify.write_once(output, result)
    print(json.dumps({"status": "pass", "rows": [{k: v for k, v in row.items() if k not in ("top_self_symbols", "input_sha256")} for row in rows]}, indent=2, sort_keys=True))
    return result


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--capture-root", type=Path, default=P / "capture")
    parser.add_argument("--baseline", type=Path, default=P.parent / "change-0886" / "self-analysis.json")
    parser.add_argument("--output", type=Path, default=P / "self-analysis.json")
    args = parser.parse_args()
    reduce_profiles(args.capture_root.resolve(), args.baseline.resolve(), args.output.resolve())
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
