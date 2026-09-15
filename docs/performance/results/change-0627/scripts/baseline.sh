#!/usr/bin/env bash
# Change 0627: first descriptive baseline of the OLE2 range-source selectors.
#
# One release binary, one CPU, sequential legs. The range-source legs are
# sleep-dominated by construction: their elapsed time is modelled, not
# measured. The owned-source control legs are CPU-bound and are run twice, in
# the same window, so the second pair is an A/A floor for the first.
set -euo pipefail

ROOT="${1:?repository root}"
OUT="${2:?output directory}"
CPU="${3:-20}"
# A staged copy, outside any Cargo target directory: a concurrent `cargo
# test`/`cargo doc` relinks the release binary in place and an in-flight run
# then dies with ENOENT. The staged copy's SHA-256 is in the packet README.
BIN="${4:?staged benchmark binary}"

# Change 0572's delayed arm inside this harness's four-parameter model:
# 1 ms of service per request, no separate overhead term, 100 MiB/s, 64 KiB.
TRANSPORT=(--range-fixed-latency-us 1000 --range-request-overhead-us 0
           --range-bandwidth-bytes-per-sec 104857600 --range-max-physical-bytes 65536)
WARMUP=20
SAMPLES=50

XLS_RANGE=xls_range_source_open,xls_range_source_open_list_worksheets,xls_range_source_open_one_cell,xls_range_source_open_all_cells,xls_range_source_open_full_text
XLS_OWNED=xls_owned_source_control_open,xls_owned_source_control_open_list_worksheets,xls_owned_source_control_open_one_cell,xls_owned_source_control_open_all_cells,xls_owned_source_control_open_full_text
PPT_RANGE=ppt_range_source_open,ppt_range_source_open_one_shape_text
PPT_OWNED=ppt_owned_source_control_open,ppt_owned_source_control_open_one_shape_text

run() {  # label cases fixture
  echo "[$(date -u +%H:%M:%S)] $1"
  taskset -c "$CPU" "$BIN" --warmup "$WARMUP" --samples "$SAMPLES" \
    --case "$2" --ole2-file "$3" "${TRANSPORT[@]}" --json "$OUT/raw/$1.json"
}

# Owned controls first (deterministic counts and the cheap leg), then the A/A
# repeat, then the sleep-bound range-source legs.
run owned-withcustomviews      "$XLS_OWNED" "$ROOT/test-data/ole/xls/WithCustomViews.xls"
run owned-conditionalformatting "$XLS_OWNED" "$ROOT/test-data/ole/xls/ConditionalFormattingSamples.xls"
run owned-54016                "$XLS_OWNED" "$ROOT/test-data/poi/test-data/spreadsheet/54016.xls"
run owned-45543                "$PPT_OWNED" "$ROOT/test-data/poi/test-data/slideshow/45543.ppt"

run aa-owned-withcustomviews      "$XLS_OWNED" "$ROOT/test-data/ole/xls/WithCustomViews.xls"
run aa-owned-conditionalformatting "$XLS_OWNED" "$ROOT/test-data/ole/xls/ConditionalFormattingSamples.xls"
run aa-owned-54016                "$XLS_OWNED" "$ROOT/test-data/poi/test-data/spreadsheet/54016.xls"
run aa-owned-45543                "$PPT_OWNED" "$ROOT/test-data/poi/test-data/slideshow/45543.ppt"

run range-45543                "$PPT_RANGE" "$ROOT/test-data/poi/test-data/slideshow/45543.ppt"
run range-withcustomviews      "$XLS_RANGE" "$ROOT/test-data/ole/xls/WithCustomViews.xls"
run range-conditionalformatting "$XLS_RANGE" "$ROOT/test-data/ole/xls/ConditionalFormattingSamples.xls"
run range-54016                "$XLS_RANGE" "$ROOT/test-data/poi/test-data/spreadsheet/54016.xls"

echo "[$(date -u +%H:%M:%S)] done"
