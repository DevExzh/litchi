#!/usr/bin/env bash
set -u
M="$1"; B="$2"; CPU="${3:-8}"
REAL=/home/zhuhe/code/litchi-worktrees/before-70d7768cc/test-data/libreoffice-core/sd/qa/unit/data/pptx/slide-section-test.pptx
echo "== 1. callgrind, public XLSX paths"
"$M/capture-cg-0653.sh" before "$B/xmlprobe.before" "$M" "$CPU" public
"$M/capture-cg-0653.sh" after  "$B/xmlprobe.after"  "$M" "$CPU" public
echo "== 2. callgrind, codec isolation pairs (re-run against the final after binary)"
"$M/capture-cg-0653.sh" before "$B/xmlprobe.before" "$M" "$CPU" mce
"$M/capture-cg-0653.sh" after  "$B/xmlprobe.after"  "$M" "$CPU" mce
echo "== 3. real-deck phases"
rm -f "$M/deck"/*.tsv
"$M/deck/run-phases.sh" "$M" "$B/probe0649.before" "$B/probe0649.after" "$REAL" "$M/deck/marker-stripped.pptx" "$CPU" 30
echo "== 4. paired XLSX timing"
rm -rf "$M/timing"
"$M/time-probe-0653.sh" "$M" "$B/xmlprobe.before" "$B/xmlprobe.after" "$CPU" 40
echo "all done"
