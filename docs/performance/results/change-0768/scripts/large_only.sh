#!/bin/sh
# 0768: harness `large` writer shape alone, before (A) and after (B), with and
# without glibc's dynamic malloc thresholds pinned. Run from the staging
# directory holding bin/A/harness and bin/B/harness at equal-length paths.
set -eu
run() { # $1 = output prefix, $2 = round, $3 = leg
  perf stat -e page-faults,instructions,cycles -x, -o "$1-perf-r$2-$3.csv" -- \
    taskset -c 6 "bin/$3/harness" --case doc_semantic_one_edit_save \
      --writer-shape large --samples 30 --warmup 5 --json "$1-r$2-$3.json" > /dev/null
}
# First observation (two sequential processes per leg, A then B, default glibc):
#   for i in 1 2; do for leg in A B; do ... --json pf-$leg-$i.json; done; done
# Pinned thresholds, then default glibc, each in ABBA order over four rounds:
for config in pinned default; do
  if [ "$config" = pinned ]; then
    export GLIBC_TUNABLES=glibc.malloc.mmap_threshold=8388608:glibc.malloc.trim_threshold=67108864
  else
    unset GLIBC_TUNABLES
  fi
  for r in 0 1 2 3; do
    if [ $((r % 2)) -eq 0 ]; then order="A B"; else order="B A"; fi
    for leg in $order; do run "$config" "$r" "$leg"; done
  done
done
