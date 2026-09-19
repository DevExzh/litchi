#!/usr/bin/env bash
# Same-binary A/A floor for the selected paragraph-index cases.
set -euo pipefail

BIN=${BIN:?set BIN to the release probe binary}
OUT=${OUT:?set OUT to a retained measurement directory}
ROOT=${ROOT:-$(git rev-parse --show-toplevel)}
CPU=${CPU:-}
PHASE=${PHASE:-before}
SAMPLES=${SAMPLES:-8}
mkdir -p "$OUT"
: > "$OUT/aa.jsonl"

run_case() {
  local route="$1" mode="$2" repetitions="$3" corpus="$4" leg="$5" pair="$6"
  local raw
  if [[ -n "$CPU" ]]; then
    raw=$(taskset -c "$CPU" "$BIN" "$route" "$mode" "$repetitions" "$corpus")
  else
    raw=$("$BIN" "$route" "$mode" "$repetitions" "$corpus")
  fi
  jq -c --arg phase "$PHASE" --arg leg "$leg" --arg pair "$pair" \
    '. + {phase: $phase, leg: $leg, pair: ($pair | tonumber)}' \
    <<<"$raw" >> "$OUT/aa.jsonl"
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
  # The real-fixture path contains no spaces in the retained corpus.
  read -r route mode repetitions corpus <<<"$spec"
  for ((sample = 0; sample < SAMPLES; sample++)); do
    run_case "$route" "$mode" "$repetitions" "$corpus" A1 "$pair"
    run_case "$route" "$mode" "$repetitions" "$corpus" A2 "$pair"
    pair=$((pair + 1))
  done
done

sha256sum "$OUT/aa.jsonl" > "$OUT/aa.jsonl.sha256"
printf 'wrote %s (%s rows)\n' "$OUT/aa.jsonl" "$(wc -l < "$OUT/aa.jsonl")"
