#!/usr/bin/env bash
# wait_quiet.sh <log> [deadline-seconds]
# Polls until CPU 17 is >=95% idle AND the 1-minute load average is < 4.0,
# or the deadline expires. Every poll is appended to <log> as JSON lines so
# the wait itself is evidence, not a claim.
set -uo pipefail
LOG=$1
DEADLINE=${2:-2400}
CPU=${CPU:-17}
IDLE_MIN=${IDLE_MIN:-95.0}
LOAD_MAX=${LOAD_MAX:-4.0}
START=$(date +%s)
: > "$LOG"
while :; do
  NOW=$(date +%s); ELAPSED=$((NOW-START))
  IDLE=$(mpstat -P "$CPU" 1 5 2>/dev/null | awk -v c="$CPU" '$1=="Average:" && $2==c {print $NF}')
  IDLE=${IDLE:-0}
  read -r L1 L5 L15 _ < /proc/loadavg
  OK=$(awk -v i="$IDLE" -v l="$L1" -v im="$IDLE_MIN" -v lm="$LOAD_MAX" \
       'BEGIN{print (i+0 >= im+0 && l+0 < lm+0) ? 1 : 0}')
  printf '{"utc":"%s","elapsed_s":%d,"cpu%s_idle_percent":%s,"load1":%s,"load5":%s,"load15":%s,"quiet":%s}\n' \
    "$(date -u +%Y-%m-%dT%H:%M:%SZ)" "$ELAPSED" "$CPU" "$IDLE" "$L1" "$L5" "$L15" "$OK" >> "$LOG"
  if [ "$OK" = "1" ]; then echo "QUIET after ${ELAPSED}s: cpu${CPU} idle ${IDLE}%, load1 ${L1}"; exit 0; fi
  if [ "$ELAPSED" -ge "$DEADLINE" ]; then
    echo "NOT QUIET after ${ELAPSED}s (deadline ${DEADLINE}s): cpu${CPU} idle ${IDLE}%, load1 ${L1}"; exit 1
  fi
  sleep 20
done
