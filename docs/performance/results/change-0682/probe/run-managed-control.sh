#!/usr/bin/env bash
# Resolve the sub-microsecond 25-query managed same-view control with a longer loop.
set -euo pipefail
BEFORE_BIN=${BEFORE_BIN:?set BEFORE_BIN}
AFTER_BIN=${AFTER_BIN:?set AFTER_BIN; use BEFORE_BIN for A/A}
OUT=${OUT:?set OUT}
CPU=${CPU:-12}
SAMPLES=${SAMPLES:-8}
mkdir -p "$OUT"
: > "$OUT/control.jsonl"
pair=0
for corpus in generated:200 generated:10000; do
  for ((sample=0; sample<SAMPLES; sample++)); do
    for leg in A1 B1 B2 A2; do
      bin="$BEFORE_BIN"
      [[ "$leg" == B* ]] && bin="$AFTER_BIN"
      taskset -c "$CPU" "$bin" managed same-count 100000 "$corpus" |
        jq -c --arg leg "$leg" --argjson pair "$pair" \
          '. + {leg: $leg, pair: $pair}' >> "$OUT/control.jsonl"
    done
    pair=$((pair+1))
  done
done
sha256sum "$OUT/control.jsonl" > "$OUT/control.jsonl.sha256"
