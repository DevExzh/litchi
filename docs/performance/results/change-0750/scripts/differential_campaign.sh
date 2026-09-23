#!/bin/bash
# Change 0750: release-mode differential campaign of the pair audit against
# two complete source audits (0747's generator, extended by 0750 with
# namespace declarations, prefixed attributes, aliases and the newly refused
# tokens), 8 seeds x 3,000,000 cases. Debug-only cross-checks are not in a
# release build; the oracle comparison is.
set -u
cd /home/zhuhe/code/litchi-worktrees/0750-xml-audit-well-formedness-gaps
export CARGO_TARGET_DIR=/home/zhuhe/code/litchi-worktrees/targets/0750 CARGO_BUILD_JOBS=6 CARGO_PROFILE_RELEASE_LTO=false CARGO_PROFILE_RELEASE_PANIC=unwind
OUT=${OUT:-/home/zhuhe/code/litchi-worktrees/scratch/0750/campaign}
mkdir -p $OUT
cargo test --release -p xml-minifier --test source_replacement --locked --offline --no-run > $OUT/build.log 2>&1
for seed in 1 2 3 4 5 6 7 8; do
  XML_MINIFIER_REPLACEMENT_CASES=${CASES:-3000000} XML_MINIFIER_REPLACEMENT_SEED=$((seed * 104729 + 7500)) \
    taskset -c ${CORE:-24} cargo test --release -p xml-minifier --test source_replacement --locked --offline -- --nocapture the_pair \
    > $OUT/seed$seed.log 2>&1
  echo "seed $seed exit $?" >> $OUT/status.txt
done
