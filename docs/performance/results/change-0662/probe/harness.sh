#!/usr/bin/env bash
set -u
BEFORE=/home/zhuhe/code/litchi-worktrees/targets/0662-before/release/litchi-perf-baseline
AFTER=/home/zhuhe/code/litchi-worktrees/targets/0662-after/release/litchi-perf-baseline
OUT="$1"
CASES="opc_source_overlay_one_part_save,opc_source_overlay_multi_part_changed,xlsx_source_backed_cell_values_one_edit_save,xlsx_source_backed_cell_values_one_percent_edit_save,xlsx_source_backed_managed_cell_values_one_edit_save,xlsx_source_backed_managed_cell_values_one_percent_edit_save,xlsx_eager_cell_values_one_edit_save,docx_source_backed_one_edit_save,pptx_source_backed_one_edit_save,pptx_source_backed_cross_copy_media_rich_lifecycle,docx_ordinary_save_lifecycle,docx_ordinary_save_atomic_publish,xlsx_ordinary_save_lifecycle,xlsx_ordinary_save_atomic_publish,pptx_ordinary_save_lifecycle,pptx_ordinary_save_atomic_publish"
run() { taskset -c 17 "$1" --warmup 5 --samples 30 --case "$CASES" --json "$2" > /dev/null 2>"$2.err"; echo "leg $2 exit=$?"; }
run "$BEFORE" "$OUT/before-1.json"
run "$AFTER"  "$OUT/after-1.json"
run "$AFTER"  "$OUT/after-2.json"
run "$BEFORE" "$OUT/before-2.json"
