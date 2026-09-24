#!/bin/bash
# Identical probe build for one arm: only the manifest directory (dependency
# tree), the remapped source path and the target directory differ by arm.
set -u
arm=$1
case $arm in
  before) TREE=/home/zhuhe/code/litchi-worktrees/0769-before-src; TARGET=/home/zhuhe/code/litchi-worktrees/targets/0769-before ;;
  after) TREE=/home/zhuhe/code/litchi-worktrees/0769-cfb-mini-sector-open-read-agreement; TARGET=/home/zhuhe/code/litchi-worktrees/targets/0769 ;;
  *) echo "arm must be before or after" >&2; exit 2 ;;
esac
export RUSTUP_TOOLCHAIN=1.95.0 CARGO_BUILD_JOBS=6 TMPDIR=/home/zhuhe/code/litchi-worktrees/scratch/0769/tmp
export CARGO_TARGET_DIR=$TARGET/probe
export RUSTFLAGS="--remap-path-prefix=$TREE=/litchi --remap-path-prefix=/home/zhuhe/code/litchi-worktrees/scratch/0769/probe=/probe"
cd /home/zhuhe/code/litchi-worktrees/scratch/0769/probe/manifests/$arm
cargo build --release --locked --offline --bins
