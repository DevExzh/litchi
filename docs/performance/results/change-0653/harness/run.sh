#!/usr/bin/env bash
# change 0653 harness legs: change 0588's four marker-free XLSX selectors (the
# no-regression leg; the generated corpora carry no MCE markers) and change
# 0638's real-file PPTX ordinary-save rows on the 0649 deck. ABBA with an A/A
# block in the same window, one pinned process per block.
set -u
OUT="$1"; B="$2"; A="$3"; DECK="$4"; CPU="${5:-8}"
XLSX=xlsx_first_cell,xlsx_full_cell_scan,xlsx_narrow_column_range_scan,xlsx_open_owned
PPTX=pptx_real_file_ordinary_save_edit,pptx_real_file_ordinary_save_lifecycle,pptx_real_file_ordinary_save_counting_publish
xrun() { taskset -c "$CPU" "$1" --warmup 3 --samples 30 --case "$XLSX" --json "$2" > /dev/null 2>&1; }
prun() { taskset -c "$CPU" "$1" --warmup 3 --samples 20 --case "$PPTX" --ooxml-file "$DECK" --json "$2" > /dev/null 2>&1; }
for kind in x p; do
  case $kind in x) F=xrun;; p) F=prun;; esac
  $F "$B" "$OUT/$kind-before.1.json"
  $F "$A" "$OUT/$kind-after.1.json"
  $F "$A" "$OUT/$kind-after.2.json"
  $F "$B" "$OUT/$kind-before.2.json"
  $F "$B" "$OUT/$kind-floorA.json"
  $F "$B" "$OUT/$kind-floorB.json"
done
echo "harness done"
