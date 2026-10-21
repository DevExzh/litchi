#!/bin/bash
# Record 0361a build: the three iWork paired benches (record 0329), release,
# one format feature each. SRC: tree of 7138746bb7 ("base") or this branch.
set -eu
SRC=$1 TARGET=$2 LABEL=$3 BIN=$4
cd "$SRC"
for pair in keynote:keynote_bench pages:pages_bench numbers:numbers_unified_bench; do
  f=${pair%%:*}; b=${pair##*:}
  CARGO_TARGET_DIR=$TARGET cargo build --release --offline -p litchi-ops --features $f --example $b
  cp "$TARGET/release/examples/$b" "$BIN/$b.$LABEL"
done
