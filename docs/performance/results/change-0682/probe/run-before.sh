#!/usr/bin/env bash
# Reproduce the baseline matrix for change 0682.
# Every invocation emits one machine-readable JSON sample. The package is
# opened before the measured loop inside the probe; the loop is the timed and
# counted region. Set BIN, OUT, and optionally CPU (an empty CPU disables
# taskset). Use CARGO_BUILD_JOBS=2 for the release build.
set -euo pipefail

BIN=${BIN:?set BIN to the release probe binary}
OUT=${OUT:?set OUT to a retained measurement directory}
ROOT=${ROOT:-$(git rev-parse --show-toplevel)}
CPU=${CPU:-}
PHASE=${PHASE:-before}
mkdir -p "$OUT"
: > "$OUT/matrix.jsonl"

run_case() {
  local route="$1" mode="$2" repetitions="$3" corpus="$4" phase="$5" index="$6"
  local raw
  if [[ -n "$CPU" ]]; then
    raw=$(taskset -c "$CPU" "$BIN" "$route" "$mode" "$repetitions" "$corpus")
  else
    raw=$("$BIN" "$route" "$mode" "$repetitions" "$corpus")
  fi
  jq -c --arg phase "$phase" --arg index "$index" \
    '. + {phase: $phase, sample_index: ($index | tonumber)}' \
    <<<"$raw" >> "$OUT/matrix.jsonl"
}

declare -a GENERATED_CORPORA=(
  generated:200
  generated:10000
)
REAL_CORPUS="$ROOT/test-data/ooxml/docx/ComplexNumberedLists.docx"

index=0
for corpus in "${GENERATED_CORPORA[@]}"; do
  for route in eager source managed; do
    for repetitions in 1 5 25; do
      run_case "$route" fresh-document "$repetitions" "$corpus" "$PHASE" "$index"
      index=$((index + 1))
      run_case "$route" fresh-count "$repetitions" "$corpus" "$PHASE" "$index"
      index=$((index + 1))
      run_case "$route" same-count "$repetitions" "$corpus" "$PHASE" "$index"
      index=$((index + 1))
    done
    run_case "$route" first-count 1 "$corpus" "$PHASE" "$index"
    index=$((index + 1))
  done
done

# The real fixture contains markup-compatibility content that the managed
# source contract correctly refuses before semantic ownership can be charged;
# retain it for the eager/source admission lanes and keep managed evidence on
# the generated corpora.
for route in eager source; do
  for repetitions in 1 5 25; do
    run_case "$route" fresh-document "$repetitions" "$REAL_CORPUS" "$PHASE" "$index"
    index=$((index + 1))
    run_case "$route" fresh-count "$repetitions" "$REAL_CORPUS" "$PHASE" "$index"
    index=$((index + 1))
    run_case "$route" same-count "$repetitions" "$REAL_CORPUS" "$PHASE" "$index"
    index=$((index + 1))
  done
  run_case "$route" first-count 1 "$REAL_CORPUS" "$PHASE" "$index"
  index=$((index + 1))
done

# Eager/source selective-view controls. Managed selective Arc-backed views are
# intentionally refused by the public contract and are covered by focused
# tests rather than represented as successful timing rows.
for corpus in "${GENERATED_CORPORA[@]}" "$REAL_CORPUS"; do
  for route in eager source; do
    for mode in fresh-paragraph same-paragraph fresh-paragraphs same-paragraphs; do
      repetitions=1
      [[ "$mode" == same-* ]] && repetitions=5
      run_case "$route" "$mode" "$repetitions" "$corpus" "$PHASE" "$index"
      index=$((index + 1))
    done
  done
  for route in source; do
    for mode in fresh-text same-text; do
      repetitions=1
      [[ "$mode" == same-* ]] && repetitions=5
      run_case "$route" "$mode" "$repetitions" "$corpus" "$PHASE" "$index"
      index=$((index + 1))
    done
  done
done

for corpus in "${GENERATED_CORPORA[@]}"; do
  for mode in fresh-text same-text; do
    repetitions=1
    [[ "$mode" == same-* ]] && repetitions=5
    run_case managed "$mode" "$repetitions" "$corpus" "$PHASE" "$index"
    index=$((index + 1))
  done
done

sha256sum "$OUT/matrix.jsonl" > "$OUT/matrix.jsonl.sha256"
printf 'wrote %s (%s rows)\n' "$OUT/matrix.jsonl" "$(wc -l < "$OUT/matrix.jsonl")"
