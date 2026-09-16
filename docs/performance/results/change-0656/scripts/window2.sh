#!/bin/bash
# change 0656: wait for a quieter window, then repeat the paired timing and perf stat.
set -u
SCR=/tmp/claude-1001/-home-zhuhe-code-litchi/709d31e6-bf68-4968-8b5c-9a2af1c22cc8/scratchpad/agents/0656
for i in $(seq 1 180); do
  L=$(cut -d' ' -f1 /proc/loadavg | cut -d. -f1)
  echo "$(date -Is) load1=$L"
  if [ "$L" -lt 20 ]; then break; fi
  sleep 30
done
mkdir -p "$SCR/timing-w2/perf"
echo "run-window load before: $(uptime)" > "$SCR/timing-w2/load.txt"
bash "$SCR/scripts/timing.sh" "$SCR/bin/litchi-perf-baseline-before" "$SCR/bin/litchi-perf-baseline-after" "$SCR/timing-w2" 30 3
echo "run-window load after timing: $(uptime)" >> "$SCR/timing-w2/load.txt"
bash "$SCR/scripts/perfstat.sh" "$SCR/bin/litchi-perf-baseline-before" "$SCR/bin/litchi-perf-baseline-after" "$SCR/timing-w2/perf" 10 1
echo "run-window load after perf: $(uptime)" >> "$SCR/timing-w2/load.txt"
python3 "$SCR/scripts/timing_report.py" "$SCR/timing-w2" > "$SCR/timing-w2/timing-summary.txt" 2>&1
python3 "$SCR/scripts/phase_per_leg.py" "$SCR/timing-w2" > "$SCR/timing-w2/phase-per-leg.txt" 2>&1
python3 "$SCR/scripts/perf_report.py" "$SCR/timing-w2/perf" > "$SCR/timing-w2/perf/perf-summary.txt" 2>&1
echo "WINDOW 2 DONE $(date -Is)"
