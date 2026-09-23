#!/usr/bin/env bash
# Change 0744 native ABBA timing: A = base 009d515bef harness, B = branch harness.
# Every process is pinned to CPU 12; order per group is A1 B1 B2 A2 A3 B3 B4 A4.
# The primary run used A = targets/0744-before (sha256 ac48ed6b...), built by the
# after leg's exact command from an equal-length base worktree; see README.md.
set -euo pipefail
A=${A:-/home/zhuhe/code/litchi-worktrees/targets/0744-before/release/litchi-perf-baseline}
B=${B:-/home/zhuhe/code/litchi-worktrees/targets/0744/release/litchi-perf-baseline}
OUT=${OUT:-raw}
CORE=${CORE:-12}
mkdir -p "$OUT"
run() { # group leg index args...
  local group=$1 leg=$2 index=$3; shift 3
  local bin=$A; [ "$leg" = B ] && bin=$B
  taskset -c "$CORE" "$bin" "$@" --json "$OUT/$group-$leg$index.json" > /dev/null 2> "$OUT/$group-$leg$index.stderr"
}
group() { # name args...
  local name=$1; shift
  local order=(A1 B1 B2 A2 A3 B3 B4 A4)
  for slot in "${order[@]}"; do
    run "$name" "${slot:0:1}" "${slot:1:1}" "$@"
    echo "$(date +%T) $name $slot done"
  done
}
group dense --case xlsx_first_cell,xlsx_full_cell_scan,xlsx_one_cell_commit_save,xlsx_one_percent_commit_save --xlsx-shape dense-wide --samples 20 --warmup 3
group light --case xlsx_open_owned,xlsx_noop_commit_save --xlsx-shape dense-wide,medium --samples 200 --warmup 20
group medium --case xlsx_first_cell,xlsx_full_cell_scan,xlsx_one_cell_commit_save,xlsx_one_percent_commit_save --xlsx-shape medium --samples 200 --warmup 20
group controls --case xlsx_source_backed_cell_values_one_edit_save,xlsx_eager_cell_values_one_edit_save --xlsx-cell-crud-shape medium --samples 50 --warmup 5
