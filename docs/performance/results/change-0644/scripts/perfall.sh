#!/usr/bin/env bash
# Native cycles for every fixture of the 0644 corpus, one leg.
#
#   ./scripts/perfall.sh <probe-binary> [mode]   > perf/perfstat-legA.tsv
#
# Run from the packet root. The fixture lists live in counts/ beside this
# script; `mode` defaults to running doc-open over the .doc list and ppt-open
# over the .ppt list, and may be set to cfb-index to price one index parse.
set -u
BIN="${1:?usage: perfall.sh <probe-binary> [mode]}"
MODE="${2:-}"
HERE="$(cd "$(dirname "$0")/.." && pwd)"

leg() {
  local mode="$1" list="$2"
  while read -r f; do
    [ -n "$f" ] || continue
    b=$(stat -c%s "$f")
    if [ "$mode" = cfb-index ]; then
      if [ "$b" -gt 262144 ]; then n1=200; n2=1200; else n1=2000; n2=12000; fi
    else
      if [ "$b" -gt 262144 ]; then n1=20; n2=120; else n1=200; n2=1200; fi
    fi
    "$HERE/scripts/perfstat.sh" "$BIN" "$mode" "$f" $n1 $n2
  done < "$list"
}

echo -e "mode\tfixture\tbytes\tcycles_per_op\tinstructions_per_op\tcycles_per_byte\tinstructions_per_byte"
if [ -n "$MODE" ]; then
  leg "$MODE" "$HERE/counts/doc-fixtures.txt"
  leg "$MODE" "$HERE/counts/ppt-fixtures.txt"
else
  leg doc-open "$HERE/counts/doc-fixtures.txt"
  leg ppt-open "$HERE/counts/ppt-fixtures.txt"
fi
