#!/bin/bash
set +e
outdir=/tmp/litchi-goal-0409-pmu
exe=/tmp/litchi-goal-0409-pmu/cache_touch
: > "$outdir/event-matrix-commands.txt"
run_event() {
    local label=$1
    local event=$2
    local bytes=$3
    local passes=$4
    local output="$outdir/event-${label}.txt"
    printf 'perf stat --no-big-num -x, -e %q -- taskset -c 3 %q 3 %q %q\n' "$event" "$exe" "$bytes" "$passes" >> "$outdir/event-matrix-commands.txt"
    timeout 15s perf stat --no-big-num -x, -e "$event" -- taskset -c 3 "$exe" 3 "$bytes" "$passes" > "$output" 2>&1
    printf 'event=%s workload_bytes=%s workload_passes=%s exit=%s output=%s\n' "$event" "$bytes" "$passes" "$?" "$output" >> "$outdir/event-matrix-results.txt"
}
: > "$outdir/event-matrix-results.txt"
# A hot 32 KiB working set should repeatedly hit in L1 after warm-up.
for event in cpu-cycles instructions cache-references cache-misses l1-dcache-loads l1-dcache-load-misses l2_cache_req_stat.dc_access_in_l2 l2_cache_req_stat.dc_hit_in_l2 l2_cache_req_stat.ls_rd_blk_c; do
    run_event "l1-${event//[^A-Za-z0-9]/_}" "$event" 32768 10000000
done
# A 256 MiB sequential stream exceeds the reported 128 MiB LLC.
for event in cpu-cycles instructions cache-references cache-misses l1-dcache-loads l1-dcache-load-misses l2_cache_req_stat.dc_access_in_l2 l2_cache_req_stat.dc_hit_in_l2 l2_cache_req_stat.ls_rd_blk_c l3_cache_accesses l3_misses; do
    run_event "llc-${event//[^A-Za-z0-9]/_}" "$event" 268435456 8
done
