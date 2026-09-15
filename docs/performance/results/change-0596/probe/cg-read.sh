#!/usr/bin/env bash
# Isolation pairs for the post-open reads (s=100 vs s=1100; per-op = delta/1000).
leg=$1
S=/tmp/claude-1001/-home-zhuhe-code-litchi/709d31e6-bf68-4968-8b5c-9a2af1c22cc8/scratchpad/agents/0596
B=/home/zhuhe/code/litchi-worktrees/targets/0596-probe/$leg/target/release/docppt_survey
R=/home/zhuhe/code/litchi
export TMPDIR=$S/tmp RAYON_NUM_THREADS=1
mkdir -p $TMPDIR $S/cg-$leg
for entry in "readcount:count:$R/test-data/poi/test-data/document/saved-by-table.doc" \
             "readtext:text:$R/test-data/poi/test-data/document/saved-by-table.doc"; do
  IFS=: read -r stem mode path <<<"$entry"
  for n in 100 1100; do
    setarch x86_64 -R taskset -c 16 valgrind --tool=callgrind \
      --callgrind-out-file=$S/cg-$leg/$stem-s$n.out --cache-sim=no --branch-sim=no \
      $B profile-read $mode "$path" 0 $n > $S/cg-$leg/$stem-s$n.stdout 2> $S/cg-$leg/$stem-s$n.stderr
    echo "rc=$? $stem-s$n" >> $S/cg-$leg/done-read.log
  done
done
echo READDONE >> $S/cg-$leg/done-read.log
