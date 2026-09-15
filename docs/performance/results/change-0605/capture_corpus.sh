#!/usr/bin/env bash
# capture_corpus.sh <binary> <leg-name> <out-file>
#
# The corpus differential change 0605's brief asks for: identical cell values,
# errors, reads, bytes and source observations for every query on **every** XLS
# fixture, not only the three the timing matrix uses.
#
# `xls_source_attribution` exists on both legs with the same interface for
# `open`, `list` and `one-cell`, so the same script drives both and the two
# outputs are compared with `diff`. One warmup and one sample per cell: this is
# a correctness sweep, not a timing capture. Each line carries the operation's
# logical counters, the harness's source and eager semantic projections, and --
# for a fixture the reader declines -- the typed refusal, so a refusal that
# moved would show up as a changed line rather than as a missing one.
set -euo pipefail
BIN=$(readlink -f "$1")
LEG="$2"
OUT=$3
REPO=${REPO:-/home/zhuhe/code/litchi}
CPU=${CPU:-25}
SCRATCH=${SCRATCH:-/tmp/claude-1001/-home-zhuhe-code-litchi/709d31e6-bf68-4968-8b5c-9a2af1c22cc8/scratchpad/agents/0605}
export TMPDIR="$SCRATCH/tmp" RAYON_NUM_THREADS=1 OMP_NUM_THREADS=1
mkdir -p "$TMPDIR"

: > "$OUT"
while IFS= read -r path; do
  stem=${path#"$REPO"/}
  for op in open list one-cell; do
    for sheet in 0 1; do
      if out=$(setarch x86_64 -R taskset -c "$CPU" "$BIN" \
          --input "$path" --mode owned-readat --operation "$op" \
          --worksheet-index "$sheet" --row 1 --column 0 \
          --warmups 1 --samples 1 2>&1); then
        printf '%s\t%s\t%s\t%s\n' "$stem" "$op" "$sheet" \
          "$(printf '%s' "$out" | python3 -c '
import json, sys
d = json.load(sys.stdin)
m = d["records"][0]["metrics"]
o = d["semantic_oracle"]
# The after leg reports two fields the before leg schema did not have; drop
# every null so the two legs are compared on what both actually reported.
drop = lambda value: {k: v for k, v in value.items() if v is not None}
print(json.dumps({
  "counters": {k: m[k] for k in
    ("read_calls","read_bytes","version_calls","len_calls","seek_calls")},
  "observation": drop(d["records"][0]["observation"]),
  "source": drop(o["source_implementation_projection"]),
  "eager": drop(o["eager_implementation_projection"]),
}, sort_keys=True))')" >> "$OUT"
      else
        printf '%s\t%s\t%s\trefused\t%s\n' "$stem" "$op" "$sheet" \
          "$(printf '%s' "$out" | tr '\n' ' ' | sed 's/  */ /g')" >> "$OUT"
      fi
    done
  done
  echo "swept $stem" >&2
done < <(find "$REPO/test-data" -name '*.xls' -type f | sort)
