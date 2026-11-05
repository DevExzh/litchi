#!/bin/bash
# Record 0326 builds. BASE: `git archive 16b800350b` (trunk with record 0325)
# plus this record's new crates/litchi-ops/examples/path_parse_bench.rs;
# AFTER: the merged branch tree. Build environment per the rules
# (CARGO_BUILD_JOBS=6, the memory-capped cargo shim); SHA-256 of every
# binary in binaries.sha256. Profiling examples link one at a time (five fat
# LTO links at once exceeded the 26 GB cap).
set -u
SRC=$1 TARGET=$2 BIN=$3
mkdir -p "$BIN"; cd "$SRC"
CARGO_TARGET_DIR=$TARGET cargo build --release --offline --manifest-path tools/perf-baseline/unified-query/Cargo.toml
cp "$TARGET/release/unified-query-baseline" "$BIN/unified-query"
CARGO_TARGET_DIR=$TARGET cargo build --release --offline --manifest-path tools/perf-baseline/unified-pptx-elements/Cargo.toml
cp "$TARGET/release/unified-pptx-elements-baseline" "$BIN/unified-pptx-elements"
CARGO_TARGET_DIR=$TARGET cargo build --release -p litchi-ops --no-default-features --features docx,xlsx,xlsb \
  --example docx_parts_bench --example xlsx_features_bench --example xlsb_features_bench --example path_parse_bench
for e in docx_parts_bench xlsx_features_bench xlsb_features_bench path_parse_bench; do cp "$TARGET/release/examples/$e" "$BIN/"; done
for e in pages_bench keynote_bench numbers_unified_bench rtf_batch_set_bench rtf_layout_bench; do
  CARGO_TARGET_DIR=$TARGET cargo build --profile profiling -p litchi-ops --no-default-features --features pages,keynote,numbers,rtf --example $e
  cp "$TARGET/profiling/examples/$e" "$BIN/"
done
