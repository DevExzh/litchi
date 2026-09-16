#!/usr/bin/env bash
# Callgrind isolation pairs for change 0637.
#
# One iteration of the probe is one operation, with the package loaded once
# outside the loop. Each pair profiles N and N+M iterations of the same
# operation; the difference divided by M is the per-operation Ir, with process
# start-up, package load and the probe's own setup cancelled out.
#
# Instruction counts rank work, not latency.
set -u
S="$1"; CPU=13
SHAPES=/home/zhuhe/code/litchi/test-data/ooxml/pptx/shapes.pptx
DECK="$S/fixtures/harness-200-slide.pptx"
mkdir -p "$S/cg"

run() { # run <tag> <binary> <file> <op> <iters>
  local tag="$1" bin="$2" file="$3" op="$4" iters="$5"
  taskset -c $CPU valgrind --tool=callgrind --callgrind-out-file="$S/cg/$tag.out" \
    --cache-sim=no --branch-sim=no "$bin" counts "$file" "$op" "$iters" \
    >/dev/null 2>"$S/cg/$tag.err"
  printf '%s\t%s\n' "$tag" "$(grep -oP 'Collected\s*:\s*\K[0-9]+' "$S/cg/$tag.err" | tail -1)"
}

for leg in "${LEGS:-before after}"; do :; done
for leg in ${LEGS:-before after}; do
  BIN="$S/bin/probe0637-pptx-$leg"
  for op in open count1 countN:8 refs1 refsN:8 slide1 slideall slides findname session text writetext slidetext; do
    tag=$(printf '%s' "$op" | tr ':' '-')
    run "$leg-shapes-$tag-4"  "$BIN" "$SHAPES" "$op" 4
    run "$leg-shapes-$tag-24" "$BIN" "$SHAPES" "$op" 24
    run "$leg-deck200-$tag-1" "$BIN" "$DECK" "$op" 1
    run "$leg-deck200-$tag-4" "$BIN" "$DECK" "$op" 4
  done
done
