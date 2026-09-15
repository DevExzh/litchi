#!/usr/bin/env bash
# Interleaved paired measurement: alternate the two legs sample by sample so
# host drift hits both equally. Each run is the selector's own measured child.
set -euo pipefail
S="$1"; N="$2"; TAG="$3"
F=/home/zhuhe/code/litchi-worktrees/targets/0594-fixtures
B=/home/zhuhe/code/litchi-worktrees/targets/0594-before/release/litchi-perf-baseline
A=/home/zhuhe/code/litchi-worktrees/targets/0594-after/release/litchi-perf-baseline
BW=/home/zhuhe/code/litchi-worktrees/before-08d968f8e
AW=/home/zhuhe/code/litchi-worktrees/0594
: > "$S/inter-$TAG-before.jsonl"; : > "$S/inter-$TAG-after.jsonl"
for i in $(seq 1 3); do
  ( cd "$BW" && taskset -c 14 "$B" --filesystem-child pptx_file_source_open "$F/pptx_file_source_open.pptx" "$F/d0.pptx" warm ) > /dev/null
  ( cd "$AW" && taskset -c 14 "$A" --filesystem-child pptx_file_source_open "$F/pptx_file_source_open.pptx" "$F/d0.pptx" warm ) > /dev/null
done
for i in $(seq 1 "$N"); do
  if [ $((i % 2)) -eq 0 ]; then
    ( cd "$BW" && taskset -c 14 "$B" --filesystem-child pptx_file_source_open "$F/pptx_file_source_open.pptx" "$F/db.pptx" warm ) >> "$S/inter-$TAG-before.jsonl"; echo >> "$S/inter-$TAG-before.jsonl"
    ( cd "$AW" && taskset -c 14 "$A" --filesystem-child pptx_file_source_open "$F/pptx_file_source_open.pptx" "$F/da.pptx" warm ) >> "$S/inter-$TAG-after.jsonl";  echo >> "$S/inter-$TAG-after.jsonl"
  else
    ( cd "$AW" && taskset -c 14 "$A" --filesystem-child pptx_file_source_open "$F/pptx_file_source_open.pptx" "$F/da.pptx" warm ) >> "$S/inter-$TAG-after.jsonl";  echo >> "$S/inter-$TAG-after.jsonl"
    ( cd "$BW" && taskset -c 14 "$B" --filesystem-child pptx_file_source_open "$F/pptx_file_source_open.pptx" "$F/db.pptx" warm ) >> "$S/inter-$TAG-before.jsonl"; echo >> "$S/inter-$TAG-before.jsonl"
  fi
done
