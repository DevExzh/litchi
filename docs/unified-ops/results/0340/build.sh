#!/bin/bash
# Record 0340 build: the record 0337 batch-closure harness, unchanged, for
# the four formats this record re-measures (Pages, Keynote, Numbers, ODC);
# release, debug=1, standalone workspace. Build environment per the rules
# (CARGO_BUILD_JOBS=6, the memory-capped cargo shim). Run once with SRC a
# tree of trunk 2083926091 ("before") and once with SRC this branch
# ("after"). SHA-256 in binaries.sha256; binaries stay outside the checkout.
set -eu
SRC=$1 TARGET=$2 BIN=$3
mkdir -p "$(dirname "$BIN")"; cd "$SRC"
CARGO_TARGET_DIR=$TARGET cargo build --release --offline --locked \
  --manifest-path tools/perf-baseline/unified-batch-closure/Cargo.toml \
  --no-default-features --features pages,keynote,numbers,odc
cp "$TARGET/release/unified-batch-closure-baseline" "$BIN"
sha256sum "$BIN"
