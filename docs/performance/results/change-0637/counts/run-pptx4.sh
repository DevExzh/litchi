#!/usr/bin/env bash
# Change 0637, PPTX-4: what a fused pass could save at most.
#
# `before` is the base: `semantic_text_from_part` runs three passes per slide
# (raw budget scan, MCE, parse). `proto` is a measurement-only checkout with the
# raw budget scan removed, so the difference is an UPPER BOUND on what fusing
# the budget checks into the parse could save; a real fusion must still
# establish those bounds somewhere inside the parse.
#
# The measured operation is `presentation.text()`, which is exactly the timed
# region of the `pptx_semantic_full_text` selector.
set -u
S="$1"; CPU=13
mkdir -p "$S/cg4"
declare -A DECKS=(
  [shapes]=/home/zhuhe/code/litchi/test-data/ooxml/pptx/shapes.pptx
  [bug62513]=/home/zhuhe/code/litchi/test-data/poi/test-data/slideshow/bug62513.pptx
  [comment45545]=/home/zhuhe/code/litchi/test-data/poi/test-data/slideshow/45545_Comment.pptx
  [deck200]="$S/fixtures/harness-200-slide.pptx"
)
run() { # run <tag> <binary> <file> <op> <iters>
  local tag="$1" bin="$2" file="$3" op="$4" iters="$5"
  taskset -c $CPU valgrind --tool=callgrind --callgrind-out-file="$S/cg4/$tag.out" \
    --cache-sim=no --branch-sim=no "$bin" counts "$file" "$op" "$iters" \
    >/dev/null 2>"$S/cg4/$tag.err"
  printf '%s\t%s\n' "$tag" "$(grep -oP 'Collected\s*:\s*\K[0-9]+' "$S/cg4/$tag.err" | tail -1)"
}
for leg in before proto; do
  BIN="$S/bin/probe0637-pptx-$leg"
  for deck in shapes bug62513 comment45545 deck200; do
    if [ "$deck" = deck200 ]; then lo=1; hi=4; else lo=4; hi=24; fi
    run "$leg-$deck-text-$lo"  "$BIN" "${DECKS[$deck]}" text "$lo"
    run "$leg-$deck-text-$hi"  "$BIN" "${DECKS[$deck]}" text "$hi"
  done
done
