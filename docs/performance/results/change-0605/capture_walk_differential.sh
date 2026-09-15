#!/usr/bin/env bash
# capture_walk_differential.sh <after-binary> <out-file>
#
# The whole-sheet walk against selected-cell queries, over every XLS fixture.
#
# `--operation all-cells` builds its oracle outside the timed region by walking
# the worksheet and then re-reading a bounded prefix of the reported positions
# one `cell_value_by_index` call at a time; the run fails if the two projections
# differ at any position. Sweeping the whole corpus with one sample therefore
# turns the harness into a corpus differential: a line per (fixture, worksheet)
# carrying the walk's cell count and digest, or the typed refusal both paths
# produce.
set -euo pipefail
BIN=$(readlink -f "$1")
OUT=$2
REPO=${REPO:-/home/zhuhe/code/litchi}
CPU=${CPU:-25}
LIMIT=${LIMIT:-128}
SCRATCH=${SCRATCH:-/tmp/claude-1001/-home-zhuhe-code-litchi/709d31e6-bf68-4968-8b5c-9a2af1c22cc8/scratchpad/agents/0605}
export TMPDIR="$SCRATCH/tmp" RAYON_NUM_THREADS=1 OMP_NUM_THREADS=1
mkdir -p "$TMPDIR"

: > "$OUT"
while IFS= read -r path; do
  stem=${path#"$REPO"/}
  for sheet in 0 1 2; do
    if out=$(setarch x86_64 -R taskset -c "$CPU" "$BIN" \
        --input "$path" --mode owned-readat --operation all-cells \
        --all-cells-strategy scan --per-cell-limit "$LIMIT" \
        --worksheet-index "$sheet" --warmups 1 --samples 1 2>&1); then
      printf '%s\t%s\tagreed\t%s\n' "$stem" "$sheet" \
        "$(printf '%s' "$out" | python3 -c '
import json, sys
d = json.load(sys.stdin)
o = d["records"][0]["observation"]
m = d["records"][0]["metrics"]
print(json.dumps({"cells": o["cells_reported"], "digest": o["outcome"],
                  "reads": m["read_calls"], "bytes": m["read_bytes"],
                  "versions": m["version_calls"]}, sort_keys=True))')" >> "$OUT"
    else
      printf '%s\t%s\trefused\t%s\n' "$stem" "$sheet" \
        "$(printf '%s' "$out" | tr '\n' ' ' | sed 's/  */ /g')" >> "$OUT"
    fi
  done
  echo "walked $stem" >&2
done < <(find "$REPO/test-data" -name '*.xls' -type f | sort)

agreed=$(grep -c $'\tagreed\t' "$OUT" || true)
disagreed=$(grep -c 'whole-sheet walk and selected-cell queries disagree' "$OUT" || true)
echo "agreed: $agreed" >> "$OUT"
echo "disagreements: $disagreed" >> "$OUT"
