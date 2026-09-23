#!/usr/bin/env python3
"""Revert one review fix at a time and record whether the targeted tests fail."""
import json, shutil, subprocess, sys
ROOT = "/home/zhuhe/code/litchi-worktrees/0744-xlsx-eager-workbook-cell-path"
X = ROOT + "/crates/litchi-xlsx/src/"
ENV = {"CARGO_TARGET_DIR": "/home/zhuhe/code/litchi-worktrees/targets/0744", "CARGO_BUILD_JOBS": "6"}
MUTATIONS = [
    ("item1-compaction-entry-local-name-only", "raw/compact.rs", [
        ("is_spreadsheetml_name(&namespace, element.name(), b\"sheetData\")", "element.local_name().as_ref() == b\"sheetData\""),
        ("lane::children_are_spreadsheetml(reader.resolver())\n", "true\n"),
        ("worksheet_root = !root_seen\n", "worksheet_root = true || !root_seen\n"),
    ], ["a_foreign_sheet_data_before_the_body_keeps_the_complete_readback", "lane_compaction_admits_only_the_worksheet_body"]),
    ("item2a-no-web-size-bound", "workbook/edit/semantic/transaction.rs", [
        ("        && after.len() <= raw::web::MAX_XML_BYTES\n", ""),
    ], ["worksheets_above_the_web_reader_limit_never_take_the_reduced_readback"]),
    ("item2b-no-collision-check", "workbook/edit/semantic/transaction.rs", [
        ("store.avoids_omitted_cells(&ranges).then_some(store)", "{ let _ = &ranges; Some(store) }"),
    ], ["a_changed_cell_reemitted_into_an_omitted_run_keeps_the_complete_readback"]),
    ("item3-no-bom-decline", "raw/worksheet/lane.rs", [
        ("        if content.starts_with(UTF8_BOM) {\n            return None;\n        }\n", ""),
    ], ["byte_order_marked_parts_always_decline"]),
    ("item3-no-name-comparison", "raw/worksheet/lane.rs", [
        ("content.get(name_start..name_end) != Some(name)\n            || ", ""),
    ], ["entries_require_the_named_tag_at_the_reader_positions"]),
    ("item6-hash-order", "raw/worksheet/edit/codec/snapshot/scan.rs", [
        ("    ordered.sort_unstable_by_key(|(_, member_indices)| member_indices.first().copied());\n", ""),
    ], ["shared_formula_refusals_name_the_first_group_in_document_order"]),
]
results = []
for name, path, edits, tests in MUTATIONS:
    full = X + path
    backup = open(full).read()
    text = backup
    for old, new in edits:
        assert old in text, (name, old)
        text = text.replace(old, new, 1)
    open(full, "w").write(text)
    try:
        command = ["cargo", "test", "-p", "litchi-xlsx", "--lib", "--locked", "--offline", "--"] + tests
        run = subprocess.run(command, cwd=ROOT, env={**__import__("os").environ, **ENV}, capture_output=True, text=True)
        summary = [line for line in run.stdout.splitlines() if line.startswith("test ")]
        results.append({"mutation": name, "file": path, "tests": tests, "exit": run.returncode,
                        "caught": run.returncode != 0, "lines": summary[-10:],
                        "compile_error": "error[" in run.stderr or "could not compile" in run.stderr})
    finally:
        open(full, "w").write(backup)
    print(json.dumps(results[-1]))
json.dump(results, open(sys.argv[1], "w"), indent=1)
