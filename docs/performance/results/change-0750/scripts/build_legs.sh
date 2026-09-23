#!/bin/bash
# Change 0750: build the harness for both legs with the identical command.
set -u
S=/home/zhuhe/code/litchi-worktrees/scratch/0750
export CARGO_BUILD_JOBS=6
( cd /home/zhuhe/code/litchi-worktrees/0750-before-src && CARGO_TARGET_DIR=/home/zhuhe/code/litchi-worktrees/targets/0750-before \
    cargo build --release --manifest-path tools/perf-baseline/Cargo.toml --locked --offline --bin litchi-perf-baseline ) > $S/build-before.log 2>&1
echo "before exit $?" >> $S/build-legs-status.txt
cp /home/zhuhe/code/litchi-worktrees/targets/0750-before/release/litchi-perf-baseline $S/bin/litchi-perf-baseline.before
( cd /home/zhuhe/code/litchi-worktrees/0750-xml-audit-well-formedness-gaps && CARGO_TARGET_DIR=/home/zhuhe/code/litchi-worktrees/targets/0750 \
    cargo build --release --manifest-path tools/perf-baseline/Cargo.toml --locked --offline --bin litchi-perf-baseline ) > $S/build-after.log 2>&1
echo "after exit $?" >> $S/build-legs-status.txt
cp /home/zhuhe/code/litchi-worktrees/targets/0750/release/litchi-perf-baseline $S/bin/litchi-perf-baseline.after
sha256sum $S/bin/litchi-perf-baseline.* >> $S/build-legs-status.txt
