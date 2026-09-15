#!/usr/bin/env bash
# Callgrind isolation pairs for the XLSB commit path.
# Per-operation instructions = (Ir at S2 samples - Ir at S1 samples) / (S2 - S1).
set -euo pipefail
SC="$1"; REPO="$2"; CPU="$3"
BEFORE=/home/zhuhe/code/litchi-worktrees/targets/0599-before/release/xlsb_crud
AFTER=/home/zhuhe/code/litchi-worktrees/0599/tools/perf-baseline/target/release/xlsb_crud
S1=5
S2=15
run() { # leg binary case fixture samples tag
  local leg="$1" bin="$2" case="$3" fixture="$4" samples="$5" tag="$6"
  local out="$SC/cg/${leg}-${case}-${tag}-${samples}.out"
  taskset -c "$CPU" valgrind --tool=callgrind --cache-sim=no --branch-sim=no \
    --callgrind-out-file="$out" \
    "$bin" --case "$case" --warmup 1 --samples "$samples" --fixture "$fixture" \
    --json /dev/null > /dev/null 2> "$SC/cg/${leg}-${case}-${tag}-${samples}.log"
  grep -m1 '^summary:' "$out" | awk '{print $2}'
}
cd "$REPO"
printf '%-12s %-45s %-10s %14s %14s %14s\n' leg case fixture "Ir@$S1" "Ir@$S2" "Ir/op"
for tag_fixture in "poi:$REPO/test-data/poi/test-data/spreadsheet/testVarious.xlsb" "syn:$SC/synthetic-4x2000x12.xlsb"; do
  tag="${tag_fixture%%:*}"; fixture="${tag_fixture#*:}"
  for case in noop_transaction_commit_save edit_one_existing_scalar_save; do
    for legbin in "before:$BEFORE" "after:$AFTER"; do
      leg="${legbin%%:*}"; bin="${legbin#*:}"
      a=$(run "$leg" "$bin" "$case" "$fixture" "$S1" "$tag")
      b=$(run "$leg" "$bin" "$case" "$fixture" "$S2" "$tag")
      per=$(( (b - a) / (S2 - S1) ))
      printf '%-12s %-45s %-10s %14d %14d %14d\n' "$leg" "$case" "$tag" "$a" "$b" "$per"
    done
  done
done
