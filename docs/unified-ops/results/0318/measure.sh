#!/bin/sh
# Record 0318: paired leaf/unified Numbers medians, base and after, B A A B.
# Build each tree's harness (release, numbers only), copy it beside this
# script as numbers_base / numbers_after, then run:
#   cargo build --release -p litchi-ops --features numbers --example numbers_unified_bench
set -e
export LITCHI_NUMBERS_BENCH_SAMPLES=8
for run in base:1 after:2 after:3 base:4; do
  binary=${run%%:*}; index=${run##*:}
  taskset -c 7 ./numbers_$binary > numbers-$index-$binary.json
done
