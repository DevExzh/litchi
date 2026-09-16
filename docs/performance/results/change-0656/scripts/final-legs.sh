#!/bin/bash
# change 0656: re-verify the after leg on the final tree's binary.
set -u
SCR=/tmp/claude-1001/-home-zhuhe-code-litchi/709d31e6-bf68-4968-8b5c-9a2af1c22cc8/scratchpad/agents/0656
REV=70d7768cc6dada420ede063f72c88dc99ad30383
mkdir -p "$SCR/counts-final" "$SCR/alloc-final" "$SCR/timing-w4/perf"
for CASE in pptx_cross_copy_plain_lifecycle pptx_cross_copy_media_rich_lifecycle; do
  for S in 1 3; do
    echo "=== after-final $CASE s$S : $(date -Is) ==="
    "$SCR/scripts/cg.sh" "$SCR/bin/litchi-perf-baseline-after-final" "after" "$CASE" "$S" "$SCR/counts-final"
  done
done
for C in plain media-rich; do
  rm -f "$SCR/alloc-final/retention-after-$C.json"
  taskset -c 11 "$SCR/bin/litchi-perf-baseline-alloc-after-final" retention --api owned --corpus "$C" \
    --samples 5 --warmup 1 --source-revision "$REV" \
    --output "$SCR/alloc-final/retention-after-$C.json" > /dev/null 2>&1
  echo "alloc after-final $C exit=$?"
  cp "$SCR/alloc/retention-before-$C.json" "$SCR/alloc-final/retention-before-$C.json"
done
echo "load before timing: $(uptime)" > "$SCR/timing-w4/load.txt"
bash "$SCR/scripts/timing.sh" "$SCR/bin/litchi-perf-baseline-before" "$SCR/bin/litchi-perf-baseline-after-final" "$SCR/timing-w4" 30 3
echo "load after timing: $(uptime)" >> "$SCR/timing-w4/load.txt"
bash "$SCR/scripts/perfstat.sh" "$SCR/bin/litchi-perf-baseline-before" "$SCR/bin/litchi-perf-baseline-after-final" "$SCR/timing-w4/perf" 10 1
echo "load after perf: $(uptime)" >> "$SCR/timing-w4/load.txt"
python3 "$SCR/scripts/timing_report.py" "$SCR/timing-w4" > "$SCR/timing-w4/timing-summary.txt" 2>&1
python3 "$SCR/scripts/phase_per_leg.py" "$SCR/timing-w4" > "$SCR/timing-w4/phase-per-leg.txt" 2>&1
python3 "$SCR/scripts/perf_report.py" "$SCR/timing-w4/perf" > "$SCR/timing-w4/perf/perf-summary.txt" 2>&1
python3 "$SCR/scripts/alloc_report.py" "$SCR" alloc-final > "$SCR/alloc-final/alloc-summary.txt" 2>&1
echo "FINAL LEGS DONE $(date -Is)"
