#!/usr/bin/env bash
# Corpus ablation driver for change 0610.
#
#   run-touch-corpus.sh <probe-binary> <checkout> <out-dir> <cpu>
#
# Runs `probe touch` for every OOXML fixture under <checkout>/test-data and
# keeps only the one-line summary per fixture (the per-part detail is kept for
# the flagship fixtures separately, because the whole-corpus detail is ~5,000
# lines of noise). `opc-reblob` runs for every format; `xlsx-hide` runs for the
# `.xlsx` fixtures, which are the only ones a semantic editor example can drive
# (0587 §4: no DOCX or PPTX example opens a real file, edits and saves).
set -u
binary="$1"; checkout="$2"; out="$3"; cpu="$4"
mkdir -p "$out"
: > "$out/touch-opc-reblob.txt"
: > "$out/touch-xlsx-hide.txt"
while IFS= read -r fixture; do
  taskset -c "$cpu" "$binary" touch "$fixture" opc-reblob 2>&1 | tail -1 \
    >> "$out/touch-opc-reblob.txt"
done < <(find "$checkout/test-data" \
  \( -name '*.xlsx' -o -name '*.xlsm' -o -name '*.xltx' -o -name '*.xltm' \
     -o -name '*.docx' -o -name '*.docm' -o -name '*.dotx' -o -name '*.dotm' \
     -o -name '*.pptx' -o -name '*.pptm' -o -name '*.potx' -o -name '*.ppsx' \
     -o -name '*.xlsb' \) | sort)
while IFS= read -r fixture; do
  taskset -c "$cpu" "$binary" touch "$fixture" xlsx-hide 2>&1 | tail -1 \
    >> "$out/touch-xlsx-hide.txt"
done < <(find "$checkout/test-data" -name '*.xlsx' | sort)
