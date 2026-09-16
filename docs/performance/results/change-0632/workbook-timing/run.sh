#!/usr/bin/env bash
# Paired ABBA timing of the source-backed open on change 0572's simulated
# transport, plus an A/A floor and a B/B floor in the same window.
# Usage: run.sh <work-dir>
set -euo pipefail
HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
WORK="${1:?usage: run.sh <work-dir>}"
BEFORE_TREE=/home/zhuhe/code/litchi-worktrees/before-c7326f680
AFTER_TREE=/home/zhuhe/code/litchi-worktrees/0632
CPU=8
WARMUP=10
SAMPLES=60
FIXTURES=(
  test-data/ooxml/xlsx/ConditionalFormattingSamples.xlsx
  test-data/ooxml/pptx/shapes.pptx
  test-data/ooxml/pptx/shape-soft-edges.pptx
  test-data/ooxml/docx/comment.docx
)
mkdir -p "$WORK"
for leg in before after; do
  case "$leg" in
    before) TREE="$BEFORE_TREE" ;;
    after)  TREE="$AFTER_TREE" ;;
  esac
  mkdir -p "$WORK/timing-$leg/src"
  cp "$HERE/src/main.rs" "$WORK/timing-$leg/src/main.rs"
  sed "s#@TREE@#$TREE#g" "$HERE/Cargo.toml.template" > "$WORK/timing-$leg/Cargo.toml"
  echo "== building $leg against $TREE"
  ( cd "$WORK/timing-$leg" && CARGO_TARGET_DIR="$WORK/target-$leg" nice -n 5 \
      cargo build --release -j 6 >/dev/null )
done
run_leg() { # <leg> <label>
  local leg="$1" label="$2"
  : > "$WORK/$label.json"
  for fixture in "${FIXTURES[@]}"; do
    ( cd "$AFTER_TREE" && taskset -c "$CPU" \
        "$WORK/target-$leg/release/workbook-range-timing" "$fixture" "$WARMUP" "$SAMPLES" ) \
      >> "$WORK/$label.json"
  done
  echo "  leg $label done"
}
for label_leg in A1:before B1:after B2:after A2:before A3:before A4:before; do
  run_leg "${label_leg#*:}" "${label_leg%%:*}"
done
echo "done: $WORK"
