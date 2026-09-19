#!/usr/bin/env bash
# Paired before/after timing. Each pair runs A1 B1 B2 A2; use the same corpus,
# CPU pin, repetitions, and process environment for both binaries.
set -euo pipefail

BEFORE_BIN=${BEFORE_BIN:?set BEFORE_BIN}
AFTER_BIN=${AFTER_BIN:?set AFTER_BIN}
OUT=${OUT:?set OUT to a retained measurement directory}
ROOT=${ROOT:-$(git rev-parse --show-toplevel)}
CPU=${CPU:-}
SAMPLES=${SAMPLES:-8}
mkdir -p "$OUT"
: > "$OUT/abba.jsonl"

run_case() {
  local bin="$1" route="$2" mode="$3" repetitions="$4" corpus="$5" leg="$6" pair="$7"
  local raw
  if [[ -n "$CPU" ]]; then
    raw=$(taskset -c "$CPU" "$bin" "$route" "$mode" "$repetitions" "$corpus")
  else
    raw=$("$bin" "$route" "$mode" "$repetitions" "$corpus")
  fi
  jq -c --arg leg "$leg" --arg pair "$pair" \
    '. + {leg: $leg, pair: ($pair | tonumber)}' \
    <<<"$raw" >> "$OUT/abba.jsonl"
}

declare -a CASES=(
  'eager fresh-count 5 generated:200'
  'eager fresh-count 5 generated:10000'
  'eager same-count 25 generated:200'
  'eager same-count 25 generated:10000'
  'source fresh-count 5 generated:200'
  'source fresh-count 5 generated:10000'
  'source same-count 25 generated:200'
  'source same-count 25 generated:10000'
  'managed fresh-count 5 generated:200'
  'managed fresh-count 5 generated:10000'
  'managed same-count 25 generated:200'
  'managed same-count 25 generated:10000'
  "eager fresh-count 5 $ROOT/test-data/ooxml/docx/ComplexNumberedLists.docx"
  "source fresh-count 5 $ROOT/test-data/ooxml/docx/ComplexNumberedLists.docx"
)

pair=0
for spec in "${CASES[@]}"; do
  read -r route mode repetitions corpus <<<"$spec"
  for ((sample = 0; sample < SAMPLES; sample++)); do
    run_case "$BEFORE_BIN" "$route" "$mode" "$repetitions" "$corpus" A1 "$pair"
    run_case "$AFTER_BIN" "$route" "$mode" "$repetitions" "$corpus" B1 "$pair"
    run_case "$AFTER_BIN" "$route" "$mode" "$repetitions" "$corpus" B2 "$pair"
    run_case "$BEFORE_BIN" "$route" "$mode" "$repetitions" "$corpus" A2 "$pair"
    pair=$((pair + 1))
  done
done

sha256sum "$OUT/abba.jsonl" > "$OUT/abba.jsonl.sha256"
printf 'wrote %s (%s rows)\n' "$OUT/abba.jsonl" "$(wc -l < "$OUT/abba.jsonl")"
