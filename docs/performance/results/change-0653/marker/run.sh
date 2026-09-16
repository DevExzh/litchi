#!/usr/bin/env bash
# change 0653 on change 0664's marker-bearing selectors: the PPTX
# opened-transaction edit and the DOCX save and full-text families, each with
# its byte-identical marker-stripped control. ABBA with two A/A blocks, one
# pinned process per block.
set -u
OUT="$1"; B="$2"; A="$3"; CPU="${4:-8}"
PPTX=pptx_marker_ordinary_save_edit,pptx_marker_control_ordinary_save_edit,pptx_marker_ordinary_save_lifecycle,pptx_marker_control_ordinary_save_lifecycle
DOCX=docx_marker_ordinary_save_edit,docx_marker_control_ordinary_save_edit,docx_marker_eager_full_text,docx_marker_control_eager_full_text,docx_marker_source_full_text,docx_marker_control_source_full_text
PTXT=pptx_marker_eager_full_text,pptx_marker_control_eager_full_text,pptx_marker_source_full_text,pptx_marker_control_source_full_text
run() { taskset -c "$CPU" "$1" --warmup 5 --samples 30 --case "$2" --json "$3" > /dev/null 2>&1; }
for kind in pptx docx ptxt; do
  case $kind in pptx) C="$PPTX";; docx) C="$DOCX";; ptxt) C="$PTXT";; esac
  run "$B" "$C" "$OUT/$kind-before.1.json"
  run "$A" "$C" "$OUT/$kind-after.1.json"
  run "$A" "$C" "$OUT/$kind-after.2.json"
  run "$B" "$C" "$OUT/$kind-before.2.json"
  run "$B" "$C" "$OUT/$kind-floorA.json"
  run "$B" "$C" "$OUT/$kind-floorB.json"
done
echo "marker selectors done"
