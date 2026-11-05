#!/bin/bash
# Record 0361 build: the three iWork unified benches (litchi-ops examples,
# record 0329) and the record 0337 batch-closure harness for the iWork
# formats, release, one format feature per bench. SRC is a tree of trunk
# f8899dd2cb ("base") or this branch ("after").
set -eu
SRC=$1 TARGET=$2 LABEL=$3 BIN=$4
cd "$SRC"
for pair in keynote:keynote_bench pages:pages_bench numbers:numbers_unified_bench; do
  f=${pair%%:*}; b=${pair##*:}
  CARGO_TARGET_DIR=$TARGET cargo build --release --offline -p litchi-ops --features $f --example $b
  cp "$TARGET/release/examples/$b" "$BIN/$b.$LABEL"
done
CARGO_TARGET_DIR=$TARGET cargo build --release --offline --locked \
  --manifest-path tools/perf-baseline/unified-batch-closure/Cargo.toml \
  --no-default-features --features pages,keynote,numbers
cp "$TARGET/release/unified-batch-closure-baseline" "$BIN/batch.$LABEL"
