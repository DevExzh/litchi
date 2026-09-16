#!/usr/bin/env bash
# Runs change 0659's whole deterministic battery for one leg.
#   run-legs.sh <probe-binary> <output-dir> <cpu>
set -euo pipefail
BIN=$1; OUT=$2; CPU=$3
HERE=$(cd "$(dirname "$0")" && pwd)
PKT=/home/zhuhe/code/litchi-worktrees/0659/docs/performance/results/change-0644
T=/home/zhuhe/code/litchi/test-data
run() { taskset -c "$CPU" "$BIN" "$@"; }
mkdir -p "$OUT"/{trace,sweeps,counts,witness}

# --- traces: the read inventory -----------------------------------------
run trace doc-open      "$T/ole/doc/documentProperties.doc" > "$OUT/trace/trace-documentProperties-doc-open.txt"
run trace doc-open      "$T/ole/doc/picture.doc"            > "$OUT/trace/trace-picture-doc-open.txt"
run trace doc-open-read "$T/ole/doc/documentProperties.doc" > "$OUT/trace/trace-documentProperties-doc-open-read.txt"
run trace ppt-open      "$T/poi/test-data/slideshow/45543.ppt" > "$OUT/trace/trace-45543-ppt-open.txt"

# --- 0644's six named sweeps, at its own witness offsets -----------------
run sweep doc-open      "$T/ole/doc/documentProperties.doc"       > "$OUT/sweeps/sweep-documentProperties-doc-open.tsv"
run sweep doc-open-read "$T/ole/doc/documentProperties.doc"       > "$OUT/sweeps/sweep-documentProperties-doc-open-read.tsv"
run sweep doc-open      "$T/ole/doc/documentProperties.doc" 8800  > "$OUT/sweeps/sweep-documentProperties-doc-open-index-region.tsv"
run sweep doc-open      "$T/ole/doc/documentProperties.doc" 520   > "$OUT/sweeps/sweep-documentProperties-doc-open-fat-region.tsv"
run sweep ppt-open      "$T/poi/test-data/slideshow/45543.ppt"        > "$OUT/sweeps/sweep-45543-ppt-open.tsv"
run sweep ppt-open      "$T/poi/test-data/slideshow/45543.ppt" 381028 > "$OUT/sweeps/sweep-45543-ppt-open-index-region.tsv"

# --- G1's corpus sweeps: 8 admitted .doc x 3 placements x 2 modes --------
while read -r FIXTURE; do
  STEM=$(basename "$FIXTURE" .doc)
  read -r _ FAT DIR _ < <(python3 "$HERE/cfboffsets.py" "$FIXTURE")
  run sweep doc-open      "$FIXTURE"       > "$OUT/sweeps/corpus-doc-open-$STEM-payload.tsv"
  run sweep doc-open      "$FIXTURE" "$FAT" > "$OUT/sweeps/corpus-doc-open-$STEM-fat.tsv"
  run sweep doc-open      "$FIXTURE" "$DIR" > "$OUT/sweeps/corpus-doc-open-$STEM-dir.tsv"
  run sweep doc-open-read "$FIXTURE"       > "$OUT/sweeps/corpus-doc-read-$STEM-payload.tsv"
  run sweep doc-open-read "$FIXTURE" "$FAT" > "$OUT/sweeps/corpus-doc-read-$STEM-fat.tsv"
  run sweep doc-open-read "$FIXTURE" "$DIR" > "$OUT/sweeps/corpus-doc-read-$STEM-dir.tsv"
done < "$PKT/counts/doc-fixtures.txt"

# --- G1's corpus sweeps: all 30 .ppt x 3 placements ---------------------
while read -r FIXTURE; do
  STEM=$(basename "$FIXTURE" .ppt)
  read -r _ FAT DIR _ < <(python3 "$HERE/cfboffsets.py" "$FIXTURE")
  run sweep ppt-open "$FIXTURE"       > "$OUT/sweeps/corpus-ppt-open-$STEM-payload.tsv"
  run sweep ppt-open "$FIXTURE" "$FAT" > "$OUT/sweeps/corpus-ppt-open-$STEM-fat.tsv"
  run sweep ppt-open "$FIXTURE" "$DIR" > "$OUT/sweeps/corpus-ppt-open-$STEM-dir.tsv"
done < "$PKT/counts/ppt-fixtures.txt"

# --- censuses ------------------------------------------------------------
run census doc-open      "$T" doc > "$OUT/counts/census-doc.tsv"
run census doc-open-read "$T" doc > "$OUT/counts/census-doc-read.tsv"
run census ppt-open      "$T" ppt > "$OUT/counts/census-ppt.tsv"

# --- 0644's four transient witnesses and the FileSource witness ----------
run transient doc-open      "$T/ole/doc/documentProperties.doc"  5 10 8703 > "$OUT/witness/transient-doc-flip5-revert10.txt"
run transient doc-open      "$T/ole/doc/documentProperties.doc"  5  9 8703 > "$OUT/witness/transient-doc-flip5-revert9.txt"
run transient doc-open-read "$T/ole/doc/documentProperties.doc" 31 36 8703 > "$OUT/witness/transient-docread-flip31-revert36.txt"
run transient doc-open-read "$T/ole/doc/documentProperties.doc" 31 35 8703 > "$OUT/witness/transient-docread-flip31-revert35.txt"
run filever "$T/ole/doc/documentProperties.doc" > "$OUT/witness/filesource-version-witness.txt"
