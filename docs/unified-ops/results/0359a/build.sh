#!/bin/bash
# Record 0359a ABBA builds: before = 7afba5f146 (record 0359 tip) with the
# old-route probe, after = the 0359a head worktree. One cargo at a time.
source /home/zhuhe/code/litchi-worktrees/.coord/w5-legacy/env.sh
P=/home/zhuhe/code/litchi-worktrees/.coord/w5-legacy/perf2
W=/home/zhuhe/code/litchi-worktrees/uo2-w5-legacy
: > $P/build.status
for leg in before after; do
  if [ $leg = before ]; then SRC=$P/before-src; else SRC=$W; fi
  (cd $SRC && CARGO_TARGET_DIR=$P/tgt-$leg cargo build --release --offline -p litchi-ops --features doc,xls,rtf --example legacy_w5_source_probe --example legacy_w5_probe) > $P/build-$leg.log 2>&1
  echo "$leg $?" >> $P/build.status
done
