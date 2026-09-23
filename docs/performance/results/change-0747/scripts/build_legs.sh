#!/bin/bash
set -u
S=/home/zhuhe/code/litchi-worktrees/scratch/0747
export CARGO_BUILD_JOBS=6
( cd /home/zhuhe/code/litchi-worktrees/0747-before-src && CARGO_TARGET_DIR=/home/zhuhe/code/litchi-worktrees/targets/0747-before \
    cargo build --release --manifest-path tools/perf-baseline/Cargo.toml --locked --offline --bin litchi-perf-baseline ) > $S/build-before.log 2>&1
echo "before exit $?" >> $S/build-legs-status.txt
cp /home/zhuhe/code/litchi-worktrees/targets/0747-before/release/litchi-perf-baseline $S/bin/litchi-perf-baseline.before-self
( cd /home/zhuhe/code/litchi-worktrees/0747-xlsx-publication-audit-reuse && CARGO_TARGET_DIR=/home/zhuhe/code/litchi-worktrees/targets/0747 \
    cargo build --release --manifest-path tools/perf-baseline/Cargo.toml --locked --offline --bin litchi-perf-baseline ) > $S/build-after2.log 2>&1
echo "after exit $?" >> $S/build-legs-status.txt
cp /home/zhuhe/code/litchi-worktrees/targets/0747/release/litchi-perf-baseline $S/bin/litchi-perf-baseline.after2
sha256sum $S/bin/litchi-perf-baseline.* >> $S/build-legs-status.txt
