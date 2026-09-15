#!/usr/bin/env bash
# Fresh callgrind isolation pairs (s=10 vs s=110; per-op = (Ir110-Ir10)/100).
S=/tmp/claude-1001/-home-zhuhe-code-litchi/709d31e6-bf68-4968-8b5c-9a2af1c22cc8/scratchpad/agents/doc-ppt
B=$S/target/release/docppt_survey
R=/home/zhuhe/code/litchi
export TMPDIR=$S/tmp RAYON_NUM_THREADS=1
mkdir -p $TMPDIR $S/cg
legs=(
 "docsmall-facade:doc-facade-open:$R/test-data/poi/test-data/document/saved-by-table.doc"
 "docbig-facade:doc-facade-open:$R/test-data/poi/test-data/document/ca.kwsymphony.www_education_School_Concert_Seat_Booking_Form_2011-12.doc"
 "docmid-facade:doc-facade-open:$R/test-data/ole/doc/FloatingPictures.doc"
 "docsnap-A:doc-snapshot-open:$R/test-data/ole/doc/HeaderFooterUnicode.doc"
 "pptmid-source:ppt-source-open:$R/test-data/poi/test-data/slideshow/45543.ppt"
 "pptmid-textedit:ppt-textedit-open:$R/test-data/poi/test-data/slideshow/45543.ppt"
)
cpu=12
for leg in "${legs[@]}"; do
  IFS=: read -r stem mode path <<<"$leg"
  for n in 10 110; do
    (
      setarch x86_64 -R taskset -c $cpu valgrind --tool=callgrind --callgrind-out-file=$S/cg/$stem-s$n.out \
        --cache-sim=no --branch-sim=no $B profile $mode "$path" 1 $n > $S/cg/$stem-s$n.stdout 2> $S/cg/$stem-s$n.stderr
      echo "rc=$? $stem-s$n" >> $S/cg/done.log
    ) &
    cpu=$((cpu+1))
  done
done
wait
echo ALLDONE >> $S/cg/done.log
