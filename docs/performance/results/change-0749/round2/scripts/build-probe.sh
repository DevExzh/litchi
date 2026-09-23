#!/bin/bash
# Identical probe build for one arm: only the manifest directory (dependency
# tree), the remapped source path and the target directory differ by arm.
set -u
arm=$1
case $arm in
  A) TREE=/home/zhuhe/code/litchi-worktrees/0749-before-src ;;
  B) TREE=/home/zhuhe/code/litchi-worktrees/0749-v1-src ;;
  C) TREE=/home/zhuhe/code/litchi-worktrees/0749-cfb-reuse-plan-validation ;;
esac
export RUSTUP_TOOLCHAIN=1.95.0 CARGO_BUILD_JOBS=6 TMPDIR=/home/zhuhe/code/litchi-worktrees/scratch/0749/tmp
export CARGO_TARGET_DIR=/home/zhuhe/code/litchi-worktrees/targets/0749/probe-$arm
export RUSTFLAGS="--remap-path-prefix=$TREE=/litchi --remap-path-prefix=/home/zhuhe/code/litchi-worktrees/scratch/0749/probe=/probe"
cd /home/zhuhe/code/litchi-worktrees/scratch/0749/probe/manifests/$arm
cargo build --release --locked --offline --bins
