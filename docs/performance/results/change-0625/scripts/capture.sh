#!/usr/bin/env bash
# Captures every measurement change 0625 cites, for one leg.
#
#   BIN=<probe binary>  REPO=<checkout the probe was built against>
#   OUT=<output dir>    LEG=before|after  CPU=19  bash capture.sh
#
# The probe is built once per leg from probe/ in this packet, with <REPO-ROOT>
# in its Cargo.toml replaced by the checkout that leg measures and its own
# CARGO_TARGET_DIR outside the repository.
set -euo pipefail
: "${BIN:?}" "${REPO:?}" "${OUT:?}" "${LEG:?}"
CPU="${CPU:-19}"
mkdir -p "$OUT/cg"

# 1. Determinism over change 0617's synthetic document, 64 in-process rebuilds
#    per storage count.
{
  echo "## $LEG"
  for n in 1 2 3 8; do
    taskset -c "$CPU" "$BIN" determinism "$n" 64
  done
} > "$OUT/determinism-$LEG.txt"

# 2. Every CFB artifact in the fixture corpus, 16 rebuilds each.
taskset -c "$CPU" "$BIN" corpus "$REPO/test-data" 16 \
  > "$OUT/corpus-$LEG.jsonl" 2> "$OUT/corpus-$LEG.err"

# 3. Callgrind isolation pairs at 2 and 12 rebuilds, three repetitions.
for rep in 1 2 3; do
  for it in 2 12; do
    for pair in "docembed:poi/test-data/document/word_with_embeded.doc" \
                "xls54016:poi/test-data/spreadsheet/54016.xls"; do
      label="${pair%%:*}"; fixture="${pair#*:}"
      taskset -c "$CPU" valgrind --tool=callgrind \
        --callgrind-out-file="$OUT/cg/$LEG-$label-$it-r$rep.out" \
        "$BIN" rebuild "$REPO/test-data/$fixture" "$it" >/dev/null 2>&1
      echo "$LEG $label $it r$rep $(grep -m1 '^summary:' "$OUT/cg/$LEG-$label-$it-r$rep.out" | awk '{print $2}')"
    done
  done
done >> "$OUT/cg-summaries.txt"
