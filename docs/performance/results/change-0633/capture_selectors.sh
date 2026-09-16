#!/usr/bin/env bash
# capture_selectors.sh <before-binary> <after-binary> <out-dir>
#
# The registered harness selectors for the XLS edit-and-save path, in the same
# A1 B1 B2 A2 order as the probe capture. `xls_numeric_source_backed_*` are the
# two registered cases that exercise `commit_source_backed`, which is the
# function this change touches; `xls_numeric_plan_only_*`,
# `xls_numeric_eager_*`, `xls_semantic_*_edit_save` and the `xls_owned_source_*`
# open selectors are the unchanged controls. Each numeric case times
# `Edit::new`, staging, commit and publication separately, with the snapshot
# open outside timing.
set -euo pipefail
BEFORE=$(readlink -f "$1")
AFTER=$(readlink -f "$2")
mkdir -p "$3"; OUT=$(readlink -f "$3")
CPU=${CPU:-9}
SCRATCH=${SCRATCH:-/tmp/claude-1001/-home-zhuhe-code-litchi/709d31e6-bf68-4968-8b5c-9a2af1c22cc8/scratchpad/agents/0633}
export TMPDIR="$SCRATCH/tmp" RAYON_NUM_THREADS=1 OMP_NUM_THREADS=1
mkdir -p "$TMPDIR"

NUMERIC=xls_numeric_eager_number_edit_save,xls_numeric_source_backed_number_edit_save,xls_numeric_eager_rk_mulrk_edit_save,xls_numeric_source_backed_rk_mulrk_edit_save,xls_numeric_plan_only_number_edit_save,xls_numeric_plan_only_rk_mulrk_edit_save
SEMANTIC=xls_semantic_open,xls_semantic_noop_edit_save,xls_semantic_one_edit_save
OWNED=xls_owned_source_open,xls_owned_source_open_list_worksheets,xls_owned_source_open_one_cell

for round in a1 b1 b2 a2; do
  case "$round" in
    a1|a2) BIN="$BEFORE" ;;
    b1|b2) BIN="$AFTER" ;;
  esac
  mkdir -p "$OUT/$round"
  setarch x86_64 -R taskset -c "$CPU" "$BIN" \
    --warmup "${WN:-10}" --samples "${SN:-100}" --case "$NUMERIC" \
    --json "$OUT/$round/numeric.json" > "$OUT/$round/numeric.stdout" 2>&1
  setarch x86_64 -R taskset -c "$CPU" "$BIN" \
    --warmup "${WS:-20}" --samples "${SS:-200}" --case "$SEMANTIC" \
    --json "$OUT/$round/semantic.json" > "$OUT/$round/semantic.stdout" 2>&1
  setarch x86_64 -R taskset -c "$CPU" "$BIN" \
    --warmup "${WS:-20}" --samples "${SS:-200}" --case "$OWNED" \
    --json "$OUT/$round/owned.json" > "$OUT/$round/owned.stdout" 2>&1
  echo "ran $round"
done
