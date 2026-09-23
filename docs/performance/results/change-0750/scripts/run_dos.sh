#!/bin/bash
# Change 0750 follow-up: time verify_source on the reviewer's inputs and the
# further adversarial namespace families, for the base (3174242282), the
# first candidate (de69fb407d, before the expanded-name fix) and the fix.
# One process per case and leg, pinned; each process runs the audit up to 5
# times, stopping after 60 s, so a slow audit runs once.
set -u
S=/home/zhuhe/code/litchi-worktrees/scratch/0750
OUT=${OUT:-$S/dos}
mkdir -p $OUT
T=/home/zhuhe/code/litchi-worktrees/targets/0750
for case_ in r1-one-tag-249990-attributes r2-one-tag-60000-attributes r3-one-tag-60000-attributes-aliased \
             r4-50000-tags-two-names f1-50000-tags-aliased f2-200-levels-redeclaring-100k-name \
             f3-100000-prefixes f4-6-aliases-of-4mb-name w1-window-in-50000-aliased-tags; do
  for leg in before prefix fixed; do
    taskset -c ${CORE:-24} $T/dos-$leg/release/litchi-dos-probe-0750 $case_ --repeat 5 --budget-seconds 60 \
      | sed "s/^{/{\"leg\":\"$leg\",/" >> $OUT/results.jsonl
    echo "$case_ $leg exit ${PIPESTATUS[0]}" >> $OUT/status.txt
  done
done
echo done >> $OUT/status.txt
