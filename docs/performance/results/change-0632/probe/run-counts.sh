#!/usr/bin/env bash
# Deterministic request, byte, allocation and source-observation counts for one
# OOXML fixture, on both legs of change 0632.
#
# Usage: run-counts.sh <work-dir>
#
# The probe source is change 0623's, verbatim.  Change 0632 removes the
# `cd-probe(46)` request that probe already classifies, so no new class is
# needed and the before leg reproduces change 0623's retained after-counts
# exactly.
set -euo pipefail
HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
WORK="${1:?usage: run-counts.sh <work-dir>}"
BEFORE_TREE=/home/zhuhe/code/litchi-worktrees/before-c7326f680
AFTER_TREE=/home/zhuhe/code/litchi-worktrees/0632
FIXTURES=(
  "test-data/ooxml/xlsx/ConditionalFormattingSamples.xlsx /xl/worksheets/sheet1.xml xlsx-132"
  "test-data/ooxml/pptx/shapes.pptx /ppt/slides/slide1.xml pptx-shapes"
  "test-data/ooxml/docx/comment.docx /word/document.xml docx-comment"
)

mkdir -p "$WORK"
for leg in before after; do
  case "$leg" in
    before) TREE="$BEFORE_TREE" ;;
    after)  TREE="$AFTER_TREE" ;;
  esac
  mkdir -p "$WORK/probe-$leg/src"
  cp "$HERE/src/main.rs" "$WORK/probe-$leg/src/main.rs"
  sed "s#@TREE@#$TREE#g" "$HERE/Cargo.toml.template" > "$WORK/probe-$leg/Cargo.toml"
  cp /home/zhuhe/code/litchi/Cargo.lock "$WORK/probe-$leg/Cargo.lock" 2>/dev/null || true
  echo "== building $leg probe against $TREE"
  ( cd "$WORK/probe-$leg" && CARGO_TARGET_DIR="$WORK/target-$leg" nice -n 5 \
      cargo build --release -j 6 >/dev/null )
  for spec in "${FIXTURES[@]}"; do
    read -r fixture part tag <<<"$spec"
    echo "== $leg $tag"
    ( cd /home/zhuhe/code/litchi-worktrees/0632 && taskset -c 8 \
        "$WORK/target-$leg/release/zip-opc-read-probe" "$fixture" "$part" ) \
      > "$WORK/counts-$tag-$leg.txt"
  done
done
echo "done: $WORK"
