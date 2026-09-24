#!/usr/bin/env bash
# Change 0765 control ABBA timing: A = base 1d1044e3ac harness, B = branch harness.
# Both binaries were built by the identical command
#   CARGO_TARGET_DIR=<target> CARGO_BUILD_JOBS=6 cargo build --release --locked --offline \
#     --manifest-path tools/perf-baseline/Cargo.toml --bin litchi-perf-baseline
# from worktrees with equal path lengths, then copied to equal-length paths
# (bin/A/..., bin/B/...) so argv[0] is the same length on both legs.
# Every process is pinned to one core; per case the order is A1 B1 B2 A2 A3 B3 B4 A4.
# `perf stat -x, -e instructions:u,cycles:u` wraps every process (whole process,
# including untimed corpus construction).
set -euo pipefail
ROOT=${ROOT:-/home/zhuhe/code/litchi-worktrees/scratch/0765}
A=${A:-$ROOT/bin/A/litchi-perf-baseline}
B=${B:-$ROOT/bin/B/litchi-perf-baseline}
OUT=${OUT:-raw}
CORE=${CORE:-20}
mkdir -p "$OUT"
run() { # group leg index args...
  local group=$1 leg=$2 index=$3; shift 3
  local bin=$A; [ "$leg" = B ] && bin=$B
  taskset -c "$CORE" perf stat -x, -e instructions:u,cycles:u -o "$OUT/$group-$leg$index.perf" \
    "$bin" "$@" --json "$OUT/$group-$leg$index.json" > /dev/null 2> "$OUT/$group-$leg$index.stderr"
}
group() { # name args...
  local name=$1; shift
  local order=(A1 B1 B2 A2 A3 B3 B4 A4)
  for slot in "${order[@]}"; do
    run "$name" "${slot:0:1}" "${slot:1:1}" "$@"
    echo "$(date +%T) $name $slot done"
  done
}
group xlsx --case xlsx_first_cell --xlsx-shape dense-wide --samples 60 --warmup 5
group pptx --case pptx_semantic_one_edit_save --semantic-shape large --samples 40 --warmup 5
group docx --case docx_semantic_one_edit_save --semantic-shape large --samples 100 --warmup 10
