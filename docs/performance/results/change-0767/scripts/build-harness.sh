#!/bin/bash
# Identical harness build for one arm; only the source tree and target dir differ.
set -u
arm=$1
case $arm in
  before) TREE=/home/zhuhe/code/litchi-worktrees/0767-before-src; TARGET=/home/zhuhe/code/litchi-worktrees/targets/0767-before ;;
  after) TREE=/home/zhuhe/code/litchi-worktrees/0767-cfb-reparse-linear; TARGET=/home/zhuhe/code/litchi-worktrees/targets/0767 ;;
esac
export RUSTUP_TOOLCHAIN=1.95.0 CARGO_BUILD_JOBS=6 TMPDIR=/home/zhuhe/code/litchi-worktrees/scratch/0767/tmp
export CARGO_TARGET_DIR=$TARGET/harness
cd $TREE
cargo build --release --locked --offline --manifest-path tools/perf-baseline/Cargo.toml --bin litchi-perf-baseline
