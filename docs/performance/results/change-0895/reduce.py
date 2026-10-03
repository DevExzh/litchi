#!/usr/bin/env python3
"""Reduce exported heaptrack stacks into disjoint owner buckets.

The stack file remains the authoritative raw artifact.  This reducer parses
only positive decimal integers, assigns each stack to its rightmost recognized
owner frame, and keeps inclusive frame shares in a separate field so they are
never added to the disjoint totals.
"""

from __future__ import annotations

import argparse
from collections import Counter
import json
from pathlib import Path
import re
from typing import Any

from common import (  # type: ignore[import-not-found]
    REDUCTION_SCHEMA,
    descriptor,
    load_plan,
    read_json,
    require,
    sha,
    write_json,
)


REQUESTED_CATEGORIES = ("legacycodec", "mce_stream", "selected_worksheet", "x14ac", "other")
OWNER_KINDS = REQUESTED_CATEGORIES
INTEGER = re.compile(r"^[0-9]+$")
PRINTED_TOTAL = re.compile(r"(?im)^calls to allocation functions:\s*([0-9][0-9,]*)\b")
UNRESOLVED_FRAME = re.compile(r"(?:^|[/:])(?:\?\?|unknown|0x[0-9a-f]+)$", re.IGNORECASE)


def _owner_kind(frame: str) -> str:
    value = frame.casefold()
    if "litchi_xlsx::raw::worksheet::x14ac::" in value:
        return "x14ac"
    if (
        "litchi_xlsx::raw::worksheet::selected::" in value
        or "sourceworksheet::stream_selection" in value
        or "sourceworksheet::select_cells" in value
        or "sourceworksheet::stream_cell" in value
    ):
        return "selected_worksheet"
    if "litchi_ooxml_common::mce::stream::" in value:
        return "mce_stream"
    if (
        "litchi_xlsx::raw::worksheet::codec::" in value
        or "litchi_ooxml_common::mce::codec::" in value
        or "process_ooxml" in value
        or "parse_source_with_observer" in value
        or "parse_with_repairs" in value
        or "parse_refusing_repairs" in value
    ):
        return "legacycodec"
    return "other"


def _category(owner_kind: str) -> str:
    return owner_kind if owner_kind in REQUESTED_CATEGORIES else "other"


def _parse_stacks(path: Path) -> tuple[Counter[str], int, Counter[str], Counter[str], Counter[str], Counter[str]]:
    stacks: Counter[str] = Counter()
    unclassified_frames: Counter[str] = Counter()
    unclassified_stacks: Counter[str] = Counter()
    unresolved_frames: Counter[str] = Counter()
    unresolved_stacks: Counter[str] = Counter()
    line_count = 0
    with path.open(encoding="utf-8") as handle:
        for line_number, raw in enumerate(handle, 1):
            text = raw.rstrip("\n")
            if not text or text.startswith("#"):
                continue
            try:
                stack, raw_cost = text.rsplit(None, 1)
            except ValueError as error:
                raise ValueError(f"stack line {line_number} has no integer cost") from error
            require(stack and INTEGER.fullmatch(raw_cost) is not None, f"stack line {line_number} does not end in an exact integer")
            cost = int(raw_cost, 10)
            require(cost > 0, f"stack line {line_number} has nonpositive allocation cost")
            # heaptrack_print's flamegraph format terminates every stack with
            # a separator.  Remove exactly that terminal empty field while
            # still rejecting empty frames in the actual call path.
            frames = stack.split(";")
            if frames and frames[-1] == "":
                frames.pop()
            require(frames and all(frames), f"stack line {line_number} contains an empty frame")
            stacks[stack] += cost
            line_count += 1
            owner = "other"
            for frame in reversed(frames):
                kind = _owner_kind(frame)
                if kind != "other":
                    owner = kind
                    break
            if owner == "other":
                unclassified_stacks[stack] += cost
            for frame in set(frames):
                if _owner_kind(frame) == "other":
                    unclassified_frames[frame] += cost
                    if UNRESOLVED_FRAME.search(frame):
                        unresolved_frames[frame] += cost
            if owner == "other" and any(UNRESOLVED_FRAME.search(frame) for frame in frames):
                unresolved_stacks[stack] += cost
    require(line_count > 0, f"stack file is empty: {path}")
    return stacks, line_count, unclassified_frames, unclassified_stacks, unresolved_frames, unresolved_stacks


def _parse_histogram(path: Path) -> tuple[int, int]:
    total = 0
    rows = 0
    with path.open(encoding="utf-8") as handle:
        for line_number, raw in enumerate(handle, 1):
            text = raw.strip()
            if not text or text.startswith("#"):
                continue
            parts = text.split()
            require(len(parts) == 2 and all(INTEGER.fullmatch(item) for item in parts), f"histogram line {line_number} is not two exact integers")
            _bytes, count = (int(item, 10) for item in parts)
            require(count > 0, f"histogram line {line_number} has nonpositive count")
            total += count
            rows += 1
    require(rows > 0, f"histogram is empty: {path}")
    return total, rows


def _printed_total(path: Path) -> tuple[int, list[int]]:
    text = path.read_text(encoding="utf-8")
    values = [int(item.replace(",", ""), 10) for item in PRINTED_TOTAL.findall(text)]
    require(values, f"calls-ranked text has no printed allocation total: {path}")
    return values[-1], values


def _stack_owner(stacks: Counter[str]) -> tuple[Counter[str], Counter[str], Counter[str]]:
    owner_totals: Counter[str] = Counter()
    category_totals: Counter[str] = Counter()
    inclusive_totals: Counter[str] = Counter()
    for stack, cost in stacks.items():
        frames = stack.split(";")
        if frames and frames[-1] == "":
            frames.pop()
        owner = "other"
        for frame in reversed(frames):
            kind = _owner_kind(frame)
            if kind != "other":
                owner = kind
                break
        owner_totals[owner] += cost
        category_totals[_category(owner)] += cost
        present = {_owner_kind(frame) for frame in frames}
        for category in REQUESTED_CATEGORIES:
            if category == "other":
                continue
            if category in present:
                inclusive_totals[category] += cost
    for category in REQUESTED_CATEGORIES:
        if category != "other":
            inclusive_totals.setdefault(category, 0)
        category_totals.setdefault(category, 0)
    for kind in OWNER_KINDS:
        owner_totals.setdefault(kind, 0)
    return owner_totals, category_totals, inclusive_totals


def _shares(values: dict[str, int], total: int) -> dict[str, dict[str, Any]]:
    return {
        key: {"allocation_cost": value, "share": (value / total if total else 0.0)}
        for key, value in values.items()
    }


def reduce_one(plan: dict[str, Any], export_root: Path, row: dict[str, Any]) -> dict[str, Any]:
    receipt_path = export_root / row["path"]
    receipt = read_json(receipt_path)
    require(receipt.get("status") == "pass" and receipt.get("exit_code") == 0, f"export receipt is not passing: {row['row']['ordinal']}")
    artifacts = receipt.get("artifacts")
    require(isinstance(artifacts, dict), "export artifacts are missing")
    paths = {}
    for name in ("stdout.txt", "stderr.txt", "stacks.txt", "histogram.txt"):
        raw = artifacts.get(name)
        require(isinstance(raw, dict), f"export artifact missing: {name}")
        path = Path(raw["path"])
        require(path.is_file() and not path.is_symlink(), f"export artifact disappeared: {path}")
        actual = descriptor(path)
        for key in ("path", "bytes", "sha256", "mode", "executable"):
            require(raw.get(key) == actual[key], f"export artifact changed: {name}")
        paths[name] = path
    stacks, line_count, unclassified_frames, unclassified_stacks, unresolved_frames, unresolved_stacks = _parse_stacks(paths["stacks.txt"])
    histogram_total, histogram_rows = _parse_histogram(paths["histogram.txt"])
    printed_total, printed_values = _printed_total(paths["stdout.txt"])
    stack_total = sum(stacks.values())
    owner_totals, category_totals, inclusive_totals = _stack_owner(stacks)
    reconciliation = {
        "collapsed_stack_total": stack_total,
        "printed_total": printed_total,
        "histogram_total": histogram_total,
        "printed_total_occurrences": printed_values,
        "histogram_rows": histogram_rows,
        "collapsed_equals_printed": stack_total == printed_total,
        "collapsed_equals_histogram": stack_total == histogram_total,
        "printed_equals_histogram": printed_total == histogram_total,
        "pass": stack_total == printed_total == histogram_total,
    }
    total = printed_total
    return {
        "row": row["row"],
        "profile": receipt["profile"],
        "export": descriptor(receipt_path),
        "calls_ranked_text": artifacts["stdout.txt"],
        "flamegraph_stacks": artifacts["stacks.txt"],
        "histogram": artifacts["histogram.txt"],
        "stack_lines": line_count,
        "unique_stacks": len(stacks),
        "reconciliation": reconciliation,
        "exclusive": _shares(dict(category_totals), total),
        "owner_kind_exclusive": _shares(dict(owner_totals), total),
        "inclusive": _shares(dict(inclusive_totals), total),
        "inclusive_is_separate_and_not_added": True,
        "unclassified_frames": [{"frame": frame, "allocation_cost": cost} for frame, cost in sorted(unclassified_frames.items())],
        "unclassified_stacks": [{"stack": stack, "allocation_cost": cost} for stack, cost in sorted(unclassified_stacks.items())],
        "unresolved_frames": [{"frame": frame, "allocation_cost": cost} for frame, cost in sorted(unresolved_frames.items())],
        "unresolved_stacks": [{"stack": stack, "allocation_cost": cost} for stack, cost in sorted(unresolved_stacks.items())],
    }


def reduce(plan_path: Path, export_root: Path, output: Path) -> dict[str, Any]:
    plan = load_plan(plan_path)
    export_root = export_root.resolve()
    export_receipt = read_json(export_root / "export.json")
    require(export_receipt.get("status") == "pass", "export receipt is not passing")
    rows = export_receipt.get("rows")
    require(rows == [{"path": f"runs/{i:02d}-{role}/receipt.json", "row": plan["rows"][i]} for i, role in enumerate(("baseline", "candidate", "candidate", "baseline"))], "export schedule changed")
    require(not output.exists() and not output.is_symlink(), f"reduction output already exists or is symlinked: {output}")
    profiles = [reduce_one(plan, export_root, row) for row in rows]
    result = {
        "schema": REDUCTION_SCHEMA,
        "status": "pass" if all(item["reconciliation"]["pass"] for item in profiles) else "fail",
        "plan": descriptor(plan_path),
        "export": descriptor(export_root / "export.json"),
        "profiles": profiles,
        "profile_count": len(profiles),
        "disjoint_categories": list(REQUESTED_CATEGORIES),
        "owner_kinds": list(OWNER_KINDS),
        "category_assignment": "rightmost/nearest recognized owner frame; generic mce_stream and selected_worksheet owners remain disjoint buckets",
        "selectedstream_alias": "selected_worksheet",
        "performance_claim": None,
    }
    write_json(output, result)
    require(result["status"] == "pass", "heaptrack stack reconciliation failed")
    return result


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--plan", type=Path, required=True)
    parser.add_argument("--export-root", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    print(json.dumps(reduce(args.plan.resolve(), args.export_root.resolve(), args.output.resolve()), indent=2, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
