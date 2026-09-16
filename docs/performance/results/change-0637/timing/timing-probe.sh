#!/usr/bin/env bash
# Paired timing of the probe operations the harness has no selector for
# (change 0637), and of the PPTX-4 prototype.
#
#   catalog legs: A = before, B = after  (the catalog memo)
#   pptx4 legs:   A = before, B = proto  (the raw-scan-free upper bound)
#
# One printed sample is the mean nanoseconds per operation over `iters`
# operations; the package is loaded once, outside the timer.
set -euo pipefail
S="$1"; CPU=13
SHAPES=/home/zhuhe/code/litchi/test-data/ooxml/pptx/shapes.pptx
BUG=/home/zhuhe/code/litchi/test-data/poi/test-data/slideshow/bug62513.pptx
DECK="$S/fixtures/harness-200-slide.pptx"
run_leg() { # run_leg <dir> <binary> <name> <file> <op> <iters> <samples>
  local dir="$1" bin="$2" name="$3" file="$4" op="$5" iters="$6" samples="$7"
  mkdir -p "$dir"
  taskset -c $CPU "$bin" time "$file" "$op" "$iters" "$samples" > "$dir/$name.txt"
}
pair() { # pair <dir> <binA> <binB> <file> <op> <iters> <samples>
  local dir="$1" ba="$2" bb="$3" file="$4" op="$5" iters="$6" samples="$7"
  run_leg "$dir" "$ba" A1 "$file" "$op" "$iters" "$samples"
  run_leg "$dir" "$bb" B1 "$file" "$op" "$iters" "$samples"
  run_leg "$dir" "$bb" B2 "$file" "$op" "$iters" "$samples"
  run_leg "$dir" "$ba" A2 "$file" "$op" "$iters" "$samples"
}
BEFORE="$S/bin/probe0637-pptx-before"
AFTER="$S/bin/probe0637-pptx-after"
PROTO="$S/bin/probe0637-pptx-proto"
date -Is >> "$S/timing/window.txt"; uptime >> "$S/timing/window.txt"
pair "$S/timing/catalog-deck200-slideall"  "$BEFORE" "$AFTER" "$DECK"   slideall  1  40
pair "$S/timing/catalog-deck200-countN8"   "$BEFORE" "$AFTER" "$DECK"   countN:8  2  40
pair "$S/timing/catalog-deck200-session"   "$BEFORE" "$AFTER" "$DECK"   session   2  40
pair "$S/timing/catalog-shapes-slideall"   "$BEFORE" "$AFTER" "$SHAPES" slideall 40  40
pair "$S/timing/catalog-shapes-session"    "$BEFORE" "$AFTER" "$SHAPES" session  40  40
pair "$S/timing/catalog-deck200-count1"    "$BEFORE" "$AFTER" "$DECK"   count1    8  40
pair "$S/timing/catalog-deck200-refs1"     "$BEFORE" "$AFTER" "$DECK"   refs1     8  40
pair "$S/timing/catalog-deck200-text"      "$BEFORE" "$AFTER" "$DECK"   text      2  40
pair "$S/timing/pptx4-deck200-text"        "$BEFORE" "$PROTO" "$DECK"   text      2  40
pair "$S/timing/pptx4-shapes-text"         "$BEFORE" "$PROTO" "$SHAPES" text     40  40
pair "$S/timing/pptx4-bug62513-text"       "$BEFORE" "$PROTO" "$BUG"    text     20  40
date -Is >> "$S/timing/window.txt"; uptime >> "$S/timing/window.txt"
