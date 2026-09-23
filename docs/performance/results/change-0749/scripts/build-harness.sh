#!/bin/bash
# Identical command for both legs; only the source tree and target dir differ.
set -u
export RUSTUP_TOOLCHAIN=1.95.0 CARGO_BUILD_JOBS=6
for leg in before after; do
  if [ $leg = before ]; then TREE=/home/zhuhe/code/litchi-worktrees/0749-before-src; TD=/home/zhuhe/code/litchi-worktrees/targets/0749-before/harness;
  else TREE=/home/zhuhe/code/litchi-worktrees/0749-cfb-reuse-plan-validation; TD=/home/zhuhe/code/litchi-worktrees/targets/0749/harness; fi
  cd $TREE
  CARGO_TARGET_DIR=$TD cargo build --release --locked --offline --manifest-path tools/perf-baseline/Cargo.toml --bin litchi-perf-baseline > /home/zhuhe/code/litchi-worktrees/scratch/0749/build-harness-$leg.log 2>&1
  echo "$leg exit=$?" >> /home/zhuhe/code/litchi-worktrees/scratch/0749/build-harness.status
done
