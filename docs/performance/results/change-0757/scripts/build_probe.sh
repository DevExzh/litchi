#!/usr/bin/env bash
# Build the probe of one leg with the identical command for both legs (change 0757).
# usage: build_probe.sh PROBE_DIR TARGET_DIR LOG
set -euo pipefail
DIR="$1"; TGT="$2"; LOG="$3"
export CARGO_TARGET_DIR="$TGT" CARGO_BUILD_JOBS=6 TMPDIR=/home/zhuhe/code/litchi-worktrees/scratch/0757/tmp
{
  cargo build --release --offline --manifest-path "$DIR/Cargo.toml"
  echo "probe build done $(date -u +%FT%TZ)"
} > "$LOG" 2>&1
