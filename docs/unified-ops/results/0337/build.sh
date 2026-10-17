#!/bin/bash
# Record 0337 build: the batch-closure harness (all 21 formats, release,
# debug=1, standalone workspace). Build environment per the rules
# (CARGO_BUILD_JOBS=6, the memory-capped cargo shim). SHA-256 in
# binaries.sha256; binaries are kept outside the checkout.
set -eu
SRC=$1 TARGET=$2 BIN=$3
mkdir -p "$BIN"; cd "$SRC"
CARGO_TARGET_DIR=$TARGET cargo build --release --offline --locked \
  --manifest-path tools/perf-baseline/unified-batch-closure/Cargo.toml
cp "$TARGET/release/unified-batch-closure-baseline" "$BIN/unified-batch-closure"
sha256sum "$BIN/unified-batch-closure"
