#!/bin/bash
# Change 0754: build the before-leg harness with the identical command used for the after leg.
set -u
S=/home/zhuhe/code/litchi-worktrees/scratch/0754
export CARGO_BUILD_JOBS=6
export TMPDIR=$S/tmp
( cd /home/zhuhe/code/litchi-worktrees/0754-before-src && CARGO_TARGET_DIR=/home/zhuhe/code/litchi-worktrees/targets/0754-before \
    cargo build --release --manifest-path tools/perf-baseline/Cargo.toml --locked --offline --bin litchi-perf-baseline ) > $S/logs/build-before.log 2>&1
echo "before exit $?" >> $S/logs/build-status.txt
