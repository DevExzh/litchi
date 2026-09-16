#!/usr/bin/env bash
# The tie-in to change 0638: its own selectors, run at this record's base.
#
# The base commit c7326f680 predates 0638 (commit 15363b3a8, harness-only), so
# 0638's diff under `tools/` is applied to the checkout before building:
#
#   git -C <checkout> show 15363b3a8 -- tools/ | git -C <checkout> apply
#
# No file under `crates/` is touched by that diff, so the crate code measured is
# the base's.
set -euo pipefail
CHECKOUT=${1:?checkout}
BIN=${2:?staged litchi-perf-baseline binary}
OUT=${3:?output directory}
CPU=${4:?cpu}
mkdir -p "${OUT}/scratch"
cd "$CHECKOUT"
taskset -c "$CPU" "$BIN" --warmup 5 --samples 20 \
  --filesystem-root "${OUT}/scratch" --json "${OUT}/0638-selectors-at-base.json" \
  --case pptx_real_file_ordinary_save_lifecycle,pptx_real_file_ordinary_save_edit,pptx_real_file_ordinary_save_counting_publish,pptx_ordinary_save_edit \
  --ooxml-file test-data/libreoffice-core/sd/qa/unit/data/pptx/slide-section-test.pptx
