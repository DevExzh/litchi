#!/bin/bash
# usage: cg-run.sh <leg> <binary> <case> <samples>
set -u
S=/tmp/claude-1001/-home-zhuhe-code-litchi/709d31e6-bf68-4968-8b5c-9a2af1c22cc8/scratchpad/agents/0655
leg=$1; bin=$2; case=$3; samples=$4
tag="${leg}-${case}-s${samples}"
out="$S/cg-$tag.out"
rm -f "$out" "$S/out/cg-$tag.json"
taskset -c 10 valgrind --tool=callgrind --callgrind-out-file="$out" \
  "$bin" --warmup 0 --samples "$samples" --case "$case" --json "$S/out/cg-$tag.json" \
  > "$S/out/cg-$tag.log" 2>&1
echo "exit=$?" >> "$S/out/cg-$tag.log"
callgrind_annotate --inclusive=yes --threshold=99.5 "$out" > "$S/out/incl-$tag.txt" 2>&1
python3 "$S/callcounts.py" "$out" package_fingerprint capture_with_provenance capture_internal load_snapshot packages_equal payload_digest part_digest compress256 project > "$S/out/calls-$tag.txt" 2>&1
tail -3 "$S/out/calls-$tag.txt"
