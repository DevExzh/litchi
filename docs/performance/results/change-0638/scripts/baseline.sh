#!/usr/bin/env bash
# First descriptive baseline for change 0638's thirty opt-in selectors.
#
# Usage: baseline.sh <repo-root> <binary> <output-dir> <cpu> <repeat-label>
#
# Six facade selectors run per (DOC, PPT) fixture pair, because `--ole2-file`
# accepts at most one file per format; the twenty-four ordinary-save selectors
# run in one invocation over the three OOXML fixtures change 0593 used.
# Every measured process is pinned to one CPU and writes into a caller-named
# --filesystem-root so the save destination is not an ambient default.
set -euo pipefail

ROOT=${1:?repo root}
BIN=${2:?binary}
OUT=${3:?output directory}
CPU=${4:?cpu}
LABEL=${5:?repeat label}

WARMUP=20
SAMPLES=50
SCRATCH="${OUT}/scratch"
mkdir -p "$OUT" "$SCRATCH"

FACADE="doc_facade_file_open,doc_facade_file_full_text,doc_facade_file_one_paragraph,\
ppt_facade_file_open,ppt_facade_file_full_text,ppt_facade_file_one_slide_text"

DOC_ONLY="doc_facade_file_open,doc_facade_file_full_text,doc_facade_file_one_paragraph"

SAVE=""
for format in docx xlsx pptx; do
  for origin in "" "real_file_"; do
    for phase in lifecycle edit atomic_publish counting_publish; do
      SAVE="${SAVE}${SAVE:+,}${format}_${origin}ordinary_save_${phase}"
    done
  done
done

run() {
  local name=$1; shift
  echo "== ${LABEL} ${name}"
  taskset -c "$CPU" "$BIN" \
    --warmup "$WARMUP" --samples "$SAMPLES" \
    --filesystem-root "$SCRATCH" \
    --json "${OUT}/${LABEL}-${name}.json" "$@"
}

cd "$ROOT"

# Facade pairs. Each run names one DOC and one PPT fixture.
run facade-small --case "$FACADE" \
  --ole2-file test-data/ole/doc/documentProperties.doc \
  --ole2-file test-data/ole/ppt/ppt_with_png.ppt

run facade-large --case "$FACADE" \
  --ole2-file test-data/ole/doc/FloatingPictures.doc \
  --ole2-file test-data/ole/ppt/SampleShow.ppt

# A DOC fixture whose facade route refuses: the refusal is the measurement.
run facade-refusal --case "$DOC_ONLY" \
  --ole2-file test-data/ole/doc/duplicate-style-names.doc

# The ordinary documented save, generated corpora and change 0593's three
# real fixtures in one invocation.
run ordinary-save --case "$SAVE" \
  --ooxml-file test-data/libreoffice-core/sw/qa/writerfilter/dmapper/data/alt-chunk-header.docx \
  --ooxml-file test-data/poi/test-data/spreadsheet/ConditionalFormattingSamples.xlsx \
  --ooxml-file test-data/libreoffice-core/sd/qa/unit/data/pptx/slide-section-test.pptx

rmdir "$SCRATCH" 2>/dev/null || true
echo "== ${LABEL} complete"
