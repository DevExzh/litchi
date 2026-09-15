#!/usr/bin/env bash
# Paired timing, CPU 16, order A1 B1 B2 A2 then an A/A pair in the same window.
set -u
S=/tmp/claude-1001/-home-zhuhe-code-litchi/709d31e6-bf68-4968-8b5c-9a2af1c22cc8/scratchpad/agents/0596
A=/home/zhuhe/code/litchi-worktrees/targets/0596-before/release/litchi-perf-baseline
B=/home/zhuhe/code/litchi-worktrees/targets/0596-after/release/litchi-perf-baseline
CASES=doc_semantic_open,doc_semantic_full_text,doc_semantic_paragraph_count,doc_semantic_one_edit_save
WARMUP=${WARMUP:-20}
SAMPLES=${SAMPLES:-300}
export TMPDIR=$S/tmp
mkdir -p $TMPDIR $S/timing2
run() { # run <binary> <outfile>
  taskset -c 16 "$1" --warmup $WARMUP --samples $SAMPLES --writer-shape tiny,large \
    --case $CASES --json "$2" > /dev/null 2> "$2.err"
  echo "rc=$? $2"
}
for leg in "A:$A:a1" "B:$B:b1" "B:$B:b2" "A:$A:a2" "A:$A:a3" "A:$A:a4"; do
  IFS=: read -r _ bin tag <<<"$leg"
  run "$bin" "$S/timing2/$tag.json"
done
echo TIMINGDONE
