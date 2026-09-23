#!/bin/bash
# Change 0750 follow-up: build both harness legs with the identical command
# (base 3174242282 and the fix), then run the release differential campaign.
set -u
S=/home/zhuhe/code/litchi-worktrees/scratch/0750
export TMPDIR=$S/tmp CARGO_BUILD_JOBS=6
mkdir -p $S/bin
( cd /home/zhuhe/code/litchi-worktrees/0750-before-src && CARGO_TARGET_DIR=/home/zhuhe/code/litchi-worktrees/targets/0750/harness-before \
    cargo build --release --manifest-path tools/perf-baseline/Cargo.toml --locked --offline --bin litchi-perf-baseline ) > $S/logs/build-before.log 2>&1
echo "build before exit $?" >> $S/followup-status.txt
( cd /home/zhuhe/code/litchi-worktrees/0750-xml-audit-well-formedness-gaps && CARGO_TARGET_DIR=/home/zhuhe/code/litchi-worktrees/targets/0750 \
    cargo build --release --manifest-path tools/perf-baseline/Cargo.toml --locked --offline --bin litchi-perf-baseline ) > $S/logs/build-after.log 2>&1
echo "build after exit $?" >> $S/followup-status.txt
cp /home/zhuhe/code/litchi-worktrees/targets/0750/harness-before/release/litchi-perf-baseline $S/bin/litchi-perf-baseline.before
cp /home/zhuhe/code/litchi-worktrees/targets/0750/release/litchi-perf-baseline $S/bin/litchi-perf-baseline.after
sha256sum $S/bin/litchi-perf-baseline.* >> $S/followup-status.txt
OUT=$S/campaign /home/zhuhe/code/litchi-worktrees/0750-xml-audit-well-formedness-gaps/docs/performance/results/change-0750/scripts/differential_campaign.sh
echo "campaign done" >> $S/followup-status.txt
