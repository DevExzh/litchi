#!/usr/bin/env bash
# Ceiling measurement for change 0604: isolation-pair `perf stat` cycles.
#
# Each case is run at S1 and S2 samples R times, pinned to one CPU; the
# per-operation figure is (median(S2) - median(S1)) / (S2 - S1), which removes
# process start-up, the fixture read and the final printing.
set -u

BIN=${BIN:?}
REPO=${REPO:?}
OUT=${OUT:?}
CPU=${CPU:-19}
S1=${S1:-10}
S2=${S2:-110}
R=${R:-11}
WARM=${WARM:-3}

DOC="$REPO/test-data/poi/test-data/document/ca.kwsymphony.www_education_School_Concert_Seat_Booking_Form_2011-12.doc"
DOCMID="$REPO/test-data/ole/doc/FloatingPictures.doc"
PPT="$REPO/test-data/poi/test-data/slideshow/45543.ppt"
XLS="$REPO/test-data/ole/xls/ConditionalFormattingSamples.xls"

run_one() { # name, samples -> "cycles instructions"
  local samples=$1; shift
  taskset -c "$CPU" perf stat -e cycles,instructions -x, -- \
    "$BIN" --warmups "$WARM" --samples "$samples" "$@" 2>&1 \
    | awk -F, '/,cycles,/{c=$1} /,instructions,/{i=$1} END{print c, i}'
}

median() { sort -n | awk '{a[NR]=$1} END{print (NR%2)? a[(NR+1)/2] : int((a[NR/2]+a[NR/2+1])/2)}'; }

measure() { # label, then probe args
  local label=$1; shift
  local lo_c=() lo_i=() hi_c=() hi_i=()
  local rep out
  for ((rep=0; rep<R; rep++)); do
    out=$(run_one "$S1" "$@"); lo_c+=("${out% *}"); lo_i+=("${out#* }")
    out=$(run_one "$S2" "$@"); hi_c+=("${out% *}"); hi_i+=("${out#* }")
  done
  local mlo_c mhi_c mlo_i mhi_i
  mlo_c=$(printf '%s\n' "${lo_c[@]}" | median)
  mhi_c=$(printf '%s\n' "${hi_c[@]}" | median)
  mlo_i=$(printf '%s\n' "${lo_i[@]}" | median)
  mhi_i=$(printf '%s\n' "${hi_i[@]}" | median)
  local per_c per_i
  per_c=$(( (mhi_c - mlo_c) / (S2 - S1) ))
  per_i=$(( (mhi_i - mlo_i) / (S2 - S1) ))
  printf '%-34s cycles/op=%-12d instructions/op=%-12d  [lo=%d hi=%d]\n' \
    "$label" "$per_c" "$per_i" "$mlo_c" "$mhi_c" | tee -a "$OUT"
  printf 'RAW %s lo_cycles=%s\nRAW %s hi_cycles=%s\n' \
    "$label" "${lo_c[*]}" "$label" "${hi_c[*]}" >> "$OUT.raw"
}

: > "$OUT"
: > "$OUT.raw"
{
  echo "# change 0604 ceiling: perf stat cycles, isolation pair S1=$S1 S2=$S2, R=$R reps, CPU $CPU"
  echo "# host: $(uname -srm), $(nproc) cores"
  echo "# binary: $BIN"
  echo "# taken: $(date -Is)"
} | tee -a "$OUT" >/dev/null

# --- opens on OWNED in-memory sources -------------------------------------
measure "doc-open/docbig"            --case doc-open        --input "$DOC"
measure "doc-open/docmid"            --case doc-open        --input "$DOCMID"
measure "ppt-open/pptmid-eager"      --case ppt-open        --input "$PPT"
measure "ppt-source-open/pptmid"     --case ppt-source-open --input "$PPT"
measure "xls-source-open/flagship"   --case xls-source-open --input "$XLS"

# --- the zero-fill term, modelled on the exact byte counts ----------------
DOCBYTES=697827,893499,4096            # 1,595,422 B, 3 eager FAT streams
DOCMIDBYTES=38775,21575,255308         # 315,658 B
PPTEAGER=311524,4096,38115             # 353,735 B
PPTSHARED=311524,4096                  # 315,620 B
XLSGLOBALS=551377                      # GlobalsBuffer::ensure, 20 fills

for spec in "docbig:$DOCBYTES" "docmid:$DOCMIDBYTES" "pptmid-eager:$PPTEAGER" \
            "pptmid-shared:$PPTSHARED" "xls-globals:$XLSGLOBALS"; do
  name=${spec%%:*}; bytes=${spec#*:}
  measure "slurp-zero/$name"   --case slurp-zero   --bytes "$bytes"
  measure "slurp-append/$name" --case slurp-append --bytes "$bytes"
  measure "memset/$name"       --case memset       --bytes "$bytes"
  measure "alloc-only/$name"   --case alloc-only   --bytes "$bytes"
done

# --- A/A floor in the same window -----------------------------------------
measure "AA-doc-open/docbig-a"       --case doc-open        --input "$DOC"
measure "AA-doc-open/docbig-b"       --case doc-open        --input "$DOC"
measure "AA-ppt-source-open/a"       --case ppt-source-open --input "$PPT"
measure "AA-ppt-source-open/b"       --case ppt-source-open --input "$PPT"
