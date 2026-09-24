#!/bin/bash
# Identical harness build for one arm; only the source tree and target dir differ.
set -u
arm=$1
case $arm in
  before) TREE=/home/zhuhe/code/litchi-worktrees/0769-before-src; TARGET=/home/zhuhe/code/litchi-worktrees/targets/0769-before ;;
  after) TREE=/home/zhuhe/code/litchi-worktrees/0769-cfb-mini-sector-open-read-agreement; TARGET=/home/zhuhe/code/litchi-worktrees/targets/0769 ;;
  *) echo "arm must be before or after" >&2; exit 2 ;;
esac
export RUSTUP_TOOLCHAIN=1.95.0 CARGO_BUILD_JOBS=6 TMPDIR=/home/zhuhe/code/litchi-worktrees/scratch/0769/tmp
export CARGO_TARGET_DIR=$TARGET/harness
export RUSTFLAGS="--remap-path-prefix=$TREE=/litchi"
cd $TREE
cargo build --release --locked --offline --manifest-path tools/perf-baseline/Cargo.toml --bin litchi-perf-baseline
