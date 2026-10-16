#!/bin/sh
# Record 0312 measurements: each bench built at 35599e1978 with its own
# format feature, run twice, pinned to core 7.
set -e
for pair in keynote:keynote_bench pages:pages_bench numbers:numbers_unified_bench; do
    cargo build --release -p litchi-ops --features "${pair%%:*}" --example "${pair##*:}"
    cp "target/release/examples/${pair##*:}" "bin/${pair##*:}"
done
for run in 1 2; do
    taskset -c 7 bin/keynote_bench --rounds 8 --iterations 30 > keynote-$run.jsonl
    taskset -c 7 bin/pages_bench --rounds 8 --iterations 30 > pages-$run.jsonl
    LITCHI_NUMBERS_BENCH_SAMPLES=8 taskset -c 7 bin/numbers_unified_bench > numbers-$run.json
done
