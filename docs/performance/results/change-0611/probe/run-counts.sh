#!/usr/bin/env bash
# Deterministic request, byte, allocation and source-observation counts for one
# OOXML fixture, on both legs of change 0611.
#
# Usage: run-counts.sh <work-dir>
#
# The probe source is change 0587's, verbatim except that `classify` names a
# read that begins at a member's local header and is longer than 640 bytes
# `member-span`. The before leg issues no such read, so its output reproduces
# change 0587's retained counts byte for byte.
set -euo pipefail
HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
WORK="${1:?usage: run-counts.sh <work-dir>}"
BEFORE_TREE=/home/zhuhe/code/litchi-worktrees/before-2d6fbeaed
AFTER_TREE=/home/zhuhe/code/litchi-worktrees/0611
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
  echo "== building $leg probe against $TREE"
  ( cd "$WORK/probe-$leg" && CARGO_TARGET_DIR="$WORK/target-$leg" nice -n 5 \
      cargo build --release -j 6 >/dev/null )
  for spec in "${FIXTURES[@]}"; do
    read -r fixture part tag <<<"$spec"
    echo "== $leg $tag"
    ( cd /home/zhuhe/code/litchi && taskset -c 31 \
        "$WORK/target-$leg/release/zip-opc-read-probe" "$fixture" "$part" ) \
      > "$WORK/counts-$tag-$leg.txt"
  done
done
echo "done: $WORK"
