#!/bin/bash
# Identical harness build for one arm; only the source tree and target dir differ.
set -u
arm=$1
case $arm in
  A) TREE=/home/zhuhe/code/litchi-worktrees/0749-before-src ;;
  C) TREE=/home/zhuhe/code/litchi-worktrees/0749-cfb-reuse-plan-validation ;;
esac
export RUSTUP_TOOLCHAIN=1.95.0 CARGO_BUILD_JOBS=6 TMPDIR=/home/zhuhe/code/litchi-worktrees/scratch/0749/tmp
export CARGO_TARGET_DIR=/home/zhuhe/code/litchi-worktrees/targets/0749/harness-$arm
cd $TREE
cargo build --release --locked --offline --manifest-path tools/perf-baseline/Cargo.toml --bin litchi-perf-baseline
