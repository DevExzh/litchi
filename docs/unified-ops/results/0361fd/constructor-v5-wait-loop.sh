#!/usr/bin/env bash
# Wait for a quiet host, then run the next protected matched-v5 attempt; repeat
# until all blocks qualify, MAX attempts are used or DEADLINE passes.
cd /home/zhuhe/code/litchi-worktrees/uo2-coordinator-constructors
V5=/deleted-coord-artifact
MAX=${MAX:-12}; DEADLINE=$(( $(date +%s) + ${MINUTES:-100} * 60 ))
busy() { ps -eo comm=,args= | grep -v -E 'v5-loop' | grep -q -E '^(cargo|rustc|rustdoc|clippy-driver|valgrind|cc|ld|clang|make|cmake|ninja|pytest|nextest|heaptrack|perf) |peculiar |run-abba|run-counting' ; }
for i in $(seq 1 $MAX); do
  quiet=0
  while [ $quiet -lt 20 ]; do
    [ $(date +%s) -gt $DEADLINE ] && { echo "deadline"; exit 0; }
    if busy; then quiet=0; else quiet=$((quiet+1)); fi; sleep 1
  done
  n=$(( $(ls -d $V5/native-a[0-9]* 2>/dev/null | grep -v '\.log$' | wc -l) + 1 ))
  echo "$(date +%T) start native-a$n"
  systemd-run --user --scope --quiet --collect -p OOMPolicy=continue -p MemoryMax=20G -p MemorySwapMax=0 -- \
    python3 docs/unified-ops/results/0361fc/constructor-abba.py \
    $V5/build-base/constructor-native $V5/build-candidate/constructor-native $V5/native-a$n \
    > $V5/native-a$n.log 2>&1
  rc=$?; echo "$(date +%T) native-a$n exit=$rc $(tail -n 1 $V5/native-a$n.log | cut -c1-200)"
  [ $rc -eq 0 ] && { echo complete; exit 0; }
  grep -q "all required blocks are already qualified" $V5/native-a$n.log && exit 0
done
