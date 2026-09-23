#!/bin/bash
set -u
cd /home/zhuhe/code/litchi-worktrees/0747-xlsx-publication-audit-reuse
export CARGO_TARGET_DIR=/home/zhuhe/code/litchi-worktrees/targets/0747 CARGO_BUILD_JOBS=6 CARGO_PROFILE_RELEASE_LTO=false CARGO_PROFILE_RELEASE_PANIC=unwind
OUT=${OUT:-/home/zhuhe/code/litchi-worktrees/scratch/0747/campaign}
cargo test --release -p xml-minifier --test source_replacement --locked --offline --no-run > $OUT/build.log 2>&1
for seed in 1 2 3 4 5 6 7 8; do
  XML_MINIFIER_REPLACEMENT_CASES=${CASES:-5000000} XML_MINIFIER_REPLACEMENT_SEED=$((seed * 7919 + 747)) \
    taskset -c 24 cargo test --release -p xml-minifier --test source_replacement --locked --offline -- --nocapture the_pair \
    > $OUT/seed$seed.log 2>&1
  echo "seed $seed exit $?" >> $OUT/status.txt
done
