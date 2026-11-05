#!/bin/sh
# Record 0314: base (75e5da7150) and after binaries, profiling builds of
# `--features pages,numbers,keynote`, pinned to core 7, B A A B.
set -e
cd "$(dirname "$0")"
mkdir -p out
for run in 1 2 3 4; do
    case $run in 1|4) b=bin; tag=base;; *) b=bin-after; tag=after;; esac
    taskset -c 7 $b/pages_bench --rounds 8 --iterations 30 > out/pages-$tag-$run.jsonl
    taskset -c 7 $b/keynote_bench --rounds 8 --iterations 30 > out/keynote-$tag-$run.jsonl
    LITCHI_NUMBERS_BENCH_SAMPLES=8 taskset -c 7 $b/numbers_unified_bench > out/numbers-$tag-$run.json
done
