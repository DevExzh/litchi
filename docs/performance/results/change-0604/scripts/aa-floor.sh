#!/usr/bin/env bash
# Wall-clock A/A floor for change 0604, stated in the program's usual terms:
# the same binary is run as leg A and leg B in A1 B1 B2 A2 order, 30 samples
# per leg, each sample the probe's own per-iteration mean over 30 iterations
# after 3 warm-ups, pinned to one CPU.
set -u
BIN=${BIN:?}
REPO=${REPO:?}
OUT=${OUT:?}
CPU=${CPU:-19}
N=${N:-30}
ITERS=${ITERS:-30}

sample() { taskset -c "$CPU" "$BIN" --warmups 3 --samples "$ITERS" "$@" \
  | sed 's/.*"elapsed_ns"://; s/}//' | awk -v n="$ITERS" '{printf "%.1f\n", $1/n}'; }

stats() { sort -n | awk '
function q(f,   i){ i=int(f*(NR-1))+1; if(i<1)i=1; if(i>NR)i=NR; return a[i] }
{a[NR]=$1}
END{ s=0; for(j=1;j<=NR;j++) s+=a[j];
  printf "n=%d p50=%.1f mean=%.1f p95=%.1f p99=%.1f\n", NR, (NR%2)?a[(NR+1)/2]:(a[NR/2]+a[NR/2+1])/2, s/NR, q(0.95), q(0.99) }'; }

: > "$OUT"
for spec in "doc-open:$REPO/test-data/poi/test-data/document/ca.kwsymphony.www_education_School_Concert_Seat_Booking_Form_2011-12.doc" \
            "ppt-source-open:$REPO/test-data/poi/test-data/slideshow/45543.ppt" \
            "xls-source-open:$REPO/test-data/ole/xls/ConditionalFormattingSamples.xls"; do
  case=${spec%%:*}; input=${spec#*:}
  A=(); B=()
  for ((i=0;i<N;i++)); do
    A+=("$(sample --case "$case" --input "$input")")
    B+=("$(sample --case "$case" --input "$input")")
    B+=("$(sample --case "$case" --input "$input")")
    A+=("$(sample --case "$case" --input "$input")")
  done
  {
    echo "## $case  (A/A, same binary both legs, ns per operation)"
    echo -n "legA: "; printf '%s\n' "${A[@]}" | stats
    echo -n "legB: "; printf '%s\n' "${B[@]}" | stats
    pa=$(printf '%s\n' "${A[@]}" | sort -n | awk '{a[NR]=$1} END{print (NR%2)?a[(NR+1)/2]:(a[NR/2]+a[NR/2+1])/2}')
    pb=$(printf '%s\n' "${B[@]}" | sort -n | awk '{a[NR]=$1} END{print (NR%2)?a[(NR+1)/2]:(a[NR/2]+a[NR/2+1])/2}')
    awk -v a="$pa" -v b="$pb" 'BEGIN{printf "floor: p50 A=%.1f B=%.1f  B/A-1 = %+.2f%%  A/B-1 = %+.2f%%\n", a, b, 100*(b/a-1), 100*(a/b-1)}'
    echo
  } | tee -a "$OUT"
done
