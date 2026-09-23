#!/usr/bin/env bash
# Build one measurement leg with the identical command for both legs (change 0757).
# usage: build_leg.sh SRC_DIR TARGET_DIR LOG
set -euo pipefail
SRC="$1"; TGT="$2"; LOG="$3"
export CARGO_TARGET_DIR="$TGT" CARGO_BUILD_JOBS=6 TMPDIR=/home/zhuhe/code/litchi-worktrees/scratch/0757/tmp
cd "$SRC"
{
  echo "src=$SRC head=$(git rev-parse HEAD) dirty=$(git status --porcelain --untracked-files=no | wc -l)"
  echo "rustc: $(rustc --version)"
  cargo build --release --locked --offline --manifest-path tools/perf-baseline/Cargo.toml --bin litchi-perf-baseline
  echo "build done $(date -u +%FT%TZ) exit=$?"
} > "$LOG" 2>&1
echo "exit=0" >> "$LOG"
