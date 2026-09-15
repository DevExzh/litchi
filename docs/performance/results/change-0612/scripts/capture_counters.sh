#!/usr/bin/env bash
# capture_counters.sh <out-file>
#
# The deterministic control for change 0612: logical read calls, read bytes and
# source-version observations per operation, asserted identical across the
# samples of each child, for both legs on all three fixtures.  `file-source`
# is the mode that counts the calls a `from_path` workbook would make;
# `owned-readat` is the in-memory mode the timing legs use, and is reported so
# the two are never confused.
set -euo pipefail
OUT=${1:?usage: capture_counters.sh <out-file>}
REPO=${REPO:-/home/zhuhe/code/litchi}
BINDIR=${BINDIR:?BINDIR must point at the leg binaries}
CPU=${CPU:-8}

cells=(
  "flagship:$REPO/test-data/ole/xls/ConditionalFormattingSamples.xls:1"
  "cv:$REPO/test-data/ole/xls/WithCustomViews.xls:1"
  "54016:$REPO/test-data/poi/test-data/spreadsheet/54016.xls:0"
)

: > "$OUT"
{
  echo "# change 0612 deterministic counters"
  echo "# host $(hostname), CPU $CPU, $(date -u +%Y-%m-%dT%H:%M:%SZ)"
  echo
  printf '%-10s %-6s %-9s %-9s %8s %10s %8s %6s\n' \
    leg fixture mode operation reads bytes version sheets
} >> "$OUT"

for cell in "${cells[@]}"; do
  IFS=: read -r stem path sheet <<<"$cell"
  for mode in file-source owned-readat; do
    for op in open list one-cell; do
      for leg in base skip; do
        "$BINDIR/leg-$leg" --input "$path" --mode "$mode" --operation "$op" \
          --worksheet-index "$sheet" --warmups 2 --samples 5 2>/dev/null \
        | CPU_LEG="$leg" STEM="$stem" MODE="$mode" OP="$op" python3 -c '
import json, os, sys
d = json.load(sys.stdin)
rows = [r["metrics"] for r in d["records"]]
first = rows[0]
for key in ("read_calls", "read_bytes", "version_calls"):
    assert all(r[key] == first[key] for r in rows), (key, [r[key] for r in rows])
proj = d["semantic_oracle"]["source_implementation_projection"]
print("%-10s %-6s %-9s %-9s %8d %10d %8d %6s" % (
    os.environ["CPU_LEG"], os.environ["STEM"], os.environ["MODE"],
    os.environ["OP"], first["read_calls"], first["read_bytes"],
    first["version_calls"], proj["worksheet_count"]))
' >> "$OUT"
      done
    done
  done
done
echo "counters -> $OUT"
