#!/bin/bash
# Identical probe build for one arm: only the manifest directory (dependency
# tree), the remapped source path and the target directory differ by arm.
set -u
arm=$1
case $arm in
  before) TREE=/home/zhuhe/code/litchi-worktrees/0767-before-src; TARGET=/home/zhuhe/code/litchi-worktrees/targets/0767-before ;;
  after) TREE=/home/zhuhe/code/litchi-worktrees/0767-cfb-reparse-linear; TARGET=/home/zhuhe/code/litchi-worktrees/targets/0767 ;;
esac
export RUSTUP_TOOLCHAIN=1.95.0 CARGO_BUILD_JOBS=6 TMPDIR=/home/zhuhe/code/litchi-worktrees/scratch/0767/tmp
export CARGO_TARGET_DIR=$TARGET/probe
export RUSTFLAGS="--remap-path-prefix=$TREE=/litchi --remap-path-prefix=/home/zhuhe/code/litchi-worktrees/scratch/0767/probe=/probe"
cd /home/zhuhe/code/litchi-worktrees/scratch/0767/probe/manifests/$arm
cargo build --release --locked --offline --bins
