#!/usr/bin/env bash
# Callgrind isolation pairs for the OLE2 DOC and PPT read paths.
# Each cell runs the same child at samples=20 and samples=320; differencing the
# two legs and dividing by 300 cancels process startup, warmup and one-time init.
set -uo pipefail

S=/tmp/claude-1001/-home-zhuhe-code-litchi/14c44904-927d-4351-97ac-5611bafb5316/scratchpad
BIN=$S/target-head/release/ole_doc_ppt_profile
OUT=$S/out-docppt
R=/home/zhuhe/code/litchi
export TMPDIR="$S/tmp" RAYON_NUM_THREADS=1 OMP_NUM_THREADS=1
mkdir -p "$TMPDIR" "$OUT/cg" "$OUT/ann" "$OUT/logs"

SMALL=170
LARGE=170

# stem:format:fixture
fixtures=(
  "docbig:doc:$R/test-data/poi/test-data/document/ca.kwsymphony.www_education_School_Concert_Seat_Booking_Form_2011-12.doc"
  "docmid:doc:$R/test-data/ole/doc/FloatingPictures.doc"
  "docsmall:doc:$R/test-data/poi/test-data/document/saved-by-table.doc"
  "pptbig:ppt:$R/test-data/office-interop/libreoffice-resaved/45543-transition-litchi.ppt"
  "pptmid:ppt:$R/test-data/poi/test-data/slideshow/45543.ppt"
  "pptsmall:ppt:$R/test-data/ole/ppt/SampleShow.ppt"
)

cpus=(2 3 4 5 6 7 8 9 10 11)
i=0
pids=()
for fx in "${fixtures[@]}"; do
  IFS=: read -r stem fmt path <<<"$fx"
  for op in open text; do
    cpu=${cpus[$((i % ${#cpus[@]}))]}
    i=$((i + 1))
    (
      for n in 170; do
        raw="$OUT/cg/$stem-$op-s$n.out"
        setarch x86_64 -R taskset -c "$cpu" \
          valgrind --tool=callgrind --callgrind-out-file="$raw" \
          --cache-sim=no --branch-sim=no \
          "$BIN" --input "$path" --format "$fmt" --operation "$op" \
          --warmups 1 --samples "$n" \
          > "$OUT/logs/$stem-$op-s$n.stdout" 2> "$OUT/logs/$stem-$op-s$n.stderr"
        rc=$?
        if [ $rc -ne 0 ]; then echo "FAIL $stem-$op-s$n rc=$rc"; exit 1; fi
        callgrind_annotate --threshold=99.9            "$raw" > "$OUT/ann/self-$stem-$op-s$n.txt"  2>/dev/null
        callgrind_annotate --threshold=99.9 --inclusive=yes "$raw" > "$OUT/ann/incl-$stem-$op-s$n.txt"  2>/dev/null
        callgrind_annotate --threshold=99.5 --tree=caller   "$raw" > "$OUT/ann/tree-$stem-$op-s$n.txt"  2>/dev/null
      done
      echo "done $stem-$op (cpu $cpu)"
    ) &
    pids+=($!)
  done
done

status=0
for pid in "${pids[@]}"; do wait "$pid" || status=1; done
echo "ALL DONE status=$status"
exit $status
