#!/usr/bin/env python3
"""Read-only replay of retained XLSX scanner costs and edit-row coverage.

No Rust execution, fresh timing, or optimization admission is implied.
Use --write once to create the canonical report; the default compares it.
"""

import argparse
import hashlib
import importlib.util
import json
from pathlib import Path
import sys

sys.dont_write_bytecode = True
HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[3]
PRIOR = HERE.parent / "change-0550"
SCAN = "litchi_xlsx::raw::worksheet::edit::codec::snapshot::scan::scan_with_limit"
SHAPES = ("medium", "dense-sparse", "noncompact", "vendor-extension")


def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def read(path):
    return json.loads(path.read_text())


def require(condition, message):
    if not condition:
        raise ValueError(message)


def inventory(shape):
    # Source-bound port of xlsx_cell_crud_inventory; source hashes are checked
    # before this function runs. Counts also bind to retained native metadata.
    result = []
    for sheet in range(4):
        if shape != "dense-sparse":
            cells = [(sheet, row, col) for row in range(48) for col in range(48)]
        elif sheet == 3:
            cells = [(sheet, index, index) for index in range(128)]
        else:
            step = (1, 4, 8)[sheet]
            cells = [(sheet, row, col) for row in range(0, 128, step)
                     for col in range(0, 128, step)]
        result.extend(cells)
    return result


def coverage():
    rows = []
    for shape in SHAPES:
        cells = inventory(shape)
        count = (len(cells) + 99) // 100
        # Source-bound port of xlsx_cell_crud_updates, not random sampling.
        updates = [cells[index * len(cells) // count] for index in range(count)]
        for repeat in (1, 2):
            native = read(PRIOR / "baseline" / f"native-r{repeat}-{shape}-c1.json")
            corpus = native["results"][0]["corpus"]
            require(corpus["entry_count"] == len(cells), "corpus cell count differs")
            require(corpus["xlsx"]["one_percent_update_count"] == len(updates),
                    "corpus update count differs")
        for case, selected in (("one-cell", cells[:1]), ("one-percent", updates)):
            touched = {(sheet, row) for sheet, row, _ in selected}
            selected_sheets = {sheet for sheet, _, _ in selected}
            planned_cells = [cell for cell in cells if cell[0] in selected_sheets]
            count_in_rows = sum((sheet, row) in touched for sheet, row, _ in cells)
            rows.append({
                "shape": shape, "case": case, "total_corpus_cells": len(cells),
                "planned_sheet_cells": len(planned_cells), "edited_cells": len(selected),
                "edited_rows": len(touched), "cells_in_edited_rows": count_in_rows,
                "cells_in_edited_rows_percent_of_planned_cells":
                    100 * count_in_rows / len(planned_cells),
                "sheets": [{
                    "sheet": sheet,
                    "cells": sum(s == sheet for s, _, _ in cells),
                    "edited_cells": sum(s == sheet for s, _, _ in selected),
                    "edited_rows": sum(s == sheet for s, _ in touched),
                    "cells_in_edited_rows": sum(s == sheet and (s, row) in touched
                                                for s, row, _ in cells),
                } for sheet in sorted(selected_sheets)],
            })
    return rows


def category(name):
    if "quick_xml::reader::" in name and "read_event_impl" in name:
        return "named_reader"
    if "quick_xml::" in name and any(part in name for part in
                                     ("resolve_event", "process_event")):
        return "named_namespace_dispatch"
    if "snapshot::scan::Scanner::" in name:
        return "named_scanner_semantics"
    return "other_direct_children"


def analyze():
    inputs = read(HERE / "inputs.json")
    for path, expected in inputs["files"].items():
        require(digest(ROOT / path) == expected, f"input changed: {path}")
    # Whole retained source inventory includes Rust, tests and harness inputs.
    sources = read(PRIOR / "baseline/source-manifest.json")
    for path, expected in sources.items():
        require(digest(ROOT / path) == expected, f"source changed: {path}")
    adrs = read(PRIOR / "adr-manifest.json")["files"]
    for path, expected in adrs.items():
        require(digest(ROOT / path) == expected, f"ADR changed: {path}")
    helper_path = PRIOR / "analyze_profiles.py"
    spec = importlib.util.spec_from_file_location("retained_0550_profiles", helper_path)
    helper = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(helper)
    report = read(PRIOR / "profile-analysis.json")
    profiles = report["stages"]["baseline"]["profiles"]
    require(len(profiles) == 8, "expected eight measured profiles")
    require({(p["shape"], p["repeat"]) for p in profiles}
            == {(shape, repeat) for shape in SHAPES for repeat in (1, 2)},
            "profile matrix differs")
    rows = []
    for profile in profiles:
        recorded = profile["measured_dump"]
        path = PRIOR / recorded["file"]
        require(digest(path) == recorded["sha256"], "raw dump changed")
        parsed = helper.parse_raw(path)
        scope = helper.classify_dump(parsed, recorded["number"])
        require(scope["role"] == "measured", "lifecycle dump selected")
        reachable, _ = helper.owner_reachable(parsed)
        selected = [key for key in reachable if parsed["functions"][key]["name"] == SCAN]
        require(len(selected) == 1, "scanner is not uniquely emitted")
        scanner_id = selected[0]
        scanner = parsed["functions"][scanner_id]
        incoming = [(key, edge) for key, fn in parsed["functions"].items()
                    for edge in fn["edges"]
                    if edge["callee_id"] == scanner_id and edge["inclusive_ir"] > 0]
        require(len(incoming) == 1 and incoming[0][0] in reachable,
                "scanner has ambiguous positive callers")
        caller_id, edge = incoming[0]
        caller_name = parsed["functions"][caller_id]["name"]
        require(caller_name in helper.NESTED_TARGETS["rewrite"], "scanner caller differs")
        children = {}
        for child_edge in scanner["edges"]:
            key = child_edge["callee_id"]
            name = parsed["functions"][key]["name"]
            row = children.setdefault(key, {"function_id": key, "name": name,
                                           "inclusive_ir": 0, "category": category(name)})
            row["inclusive_ir"] += child_edge["inclusive_ir"]
        costs = {"scanner_self": scanner["self_ir"]}
        for row in children.values():
            costs[row["category"]] = costs.get(row["category"], 0) + row["inclusive_ir"]
        require(sum(costs.values()) == edge["inclusive_ir"], "scanner partition differs")
        require(edge["inclusive_ir"] > 0, "no measured scanner work")
        rows.append({
            "shape": profile["shape"], "repeat": profile["repeat"],
            "raw_file": str(path.relative_to(ROOT)), "raw_sha256": digest(path),
            "owner": helper.OWNER, "positive_owner_parent": scope["owner_parent"],
            "rewrite": caller_name, "scanner": SCAN,
            "commit_ir": parsed["summary_ir"], "scanner_ir": edge["inclusive_ir"],
            "partition_ir": costs,
            "partition_percent_of_scanner":
                {key: 100 * value / edge["inclusive_ir"] for key, value in costs.items()},
            "children": sorted(children.values(), key=lambda row: (-row["inclusive_ir"],
                                                                      row["function_id"])),
        })
    return {
        "schema": "xlsx_layout_handoff_feasibility_0551_v1",
        "performance_claim": "none",
        "source_revision": inputs["source_revision"],
        "source_inventory_checked": len(sources), "adr_hashes_checked": len(adrs),
        "inputs_sha256": digest(HERE / "inputs.json"),
        "profile_rows": rows, "edit_row_coverage": coverage(),
        "limitations": [
            "Reanalysis of retained 0550 captures; no fresh timing, allocation or hardware run.",
            "Direct scanner partitions are disjoint; names do not separate inlined work.",
            "Named reader/namespace costs are not proven removable costs or latency fractions.",
            "Scanner semantics would still need to execute in a shared full-Layout observer.",
            "Row coverage is a source-bound corpus calculation, not executed candidate work.",
            "No proof representation, memory estimate or runtime candidate is admitted.",
        ],
    }


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--write", action="store_true")
    args = parser.parse_args()
    result = json.dumps(analyze(), indent=2, sort_keys=True) + "\n"
    target = HERE / "analysis.json"
    if args.write:
        with target.open("x") as stream:
            stream.write(result)
        print("Created analysis.json")
    else:
        require(target.read_text() == result, "canonical analysis replay differs")
        print("PASS: source/ADR custody, eight scanner partitions, eight edit-row coverage rows")
