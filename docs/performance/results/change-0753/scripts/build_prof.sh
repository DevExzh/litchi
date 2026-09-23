#!/usr/bin/env bash
# Frame-pointer + line-table build of the harness, used only for profile attribution (never timed).
# usage: build_prof.sh SRC_DIR TARGET_DIR LOG
set -euo pipefail
SRC="$1"; TGT="$2"; LOG="$3"
export CARGO_TARGET_DIR="$TGT" CARGO_BUILD_JOBS=6 TMPDIR=/home/zhuhe/code/litchi-worktrees/scratch/0753/tmp
export RUSTFLAGS="-C force-frame-pointers=yes" CARGO_PROFILE_RELEASE_DEBUG=line-tables-only
cd "$SRC"
{
  echo "src=$SRC head=$(git rev-parse HEAD)"
  cargo build --release --locked --offline --manifest-path tools/perf-baseline/Cargo.toml --bin litchi-perf-baseline
  echo "build done $(date -u +%FT%TZ)"
} > "$LOG" 2>&1
