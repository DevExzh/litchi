#!/usr/bin/env bash
# Callgrind isolation pairs (s=10 vs s=110; per-op Ir = (Ir110-Ir10)/100), CPU 16.
leg=$1
S=/tmp/claude-1001/-home-zhuhe-code-litchi/709d31e6-bf68-4968-8b5c-9a2af1c22cc8/scratchpad/agents/0596
B=/home/zhuhe/code/litchi-worktrees/targets/0596-probe/$leg/target/release/docppt_survey
R=/home/zhuhe/code/litchi
export TMPDIR=$S/tmp RAYON_NUM_THREADS=1
mkdir -p $TMPDIR $S/cg-$leg
legs=(
 "docsmall:doc-facade-open:$R/test-data/poi/test-data/document/saved-by-table.doc"
 "docmid:doc-facade-open:$R/test-data/ole/doc/FloatingPictures.doc"
 "docbig:doc-facade-open:$R/test-data/poi/test-data/document/ca.kwsymphony.www_education_School_Concert_Seat_Booking_Form_2011-12.doc"
 "docpic:doc-facade-open:$R/test-data/ole/doc/picture.doc"
)
for entry in "${legs[@]}"; do
  IFS=: read -r stem mode path <<<"$entry"
  for n in 10 110; do
    setarch x86_64 -R taskset -c 16 valgrind --tool=callgrind \
      --callgrind-out-file=$S/cg-$leg/$stem-s$n.out --cache-sim=no --branch-sim=no \
      $B profile $mode "$path" 1 $n > $S/cg-$leg/$stem-s$n.stdout 2> $S/cg-$leg/$stem-s$n.stderr
    echo "rc=$? $stem-s$n" >> $S/cg-$leg/done.log
  done
done
echo ALLDONE >> $S/cg-$leg/done.log
