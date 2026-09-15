#!/bin/bash
# usage: cg-run.sh <leg> <binary> <case> <samples>
set -u
S=/tmp/claude-1001/-home-zhuhe-code-litchi/709d31e6-bf68-4968-8b5c-9a2af1c22cc8/scratchpad/agents/0590
leg=$1; bin=$2; case=$3; samples=$4
tag="${leg}-${case}-s${samples}"
out="$S/cg-$tag.out"
rm -f "$out" "$S/cg-$tag.json"
taskset -c 10 valgrind --tool=callgrind --callgrind-out-file="$out" \
  "$bin" --warmup 0 --samples "$samples" --case "$case" --json "$S/cg-$tag.json" \
  > "$S/cg-$tag.log" 2>&1
echo "exit=$?" >> "$S/cg-$tag.log"
callgrind_annotate --inclusive=yes --threshold=99.5 "$out" > "$S/incl-$tag.txt" 2>&1
python3 "$S/callcounts.py" "$out" package_fingerprint capture_with_provenance load_snapshot packages_equal > "$S/calls-$tag.txt" 2>&1
grep -E "^(Collected|  Collected|==.*Collected)" "$S/cg-$tag.log" | tail -2
tail -3 "$S/calls-$tag.txt"
