#!/bin/bash
# Record 0337 measurement driver: three trials (processes) of each set, one
# pinned CPU, nothing else heavy running; every case is an in-process ABBA
# (5 rounds, 10 samples per leg) after an untimed equivalence check.
#   A: fixture of 10,000 targets, 100 / 1,000 / 10,000 operations.
#   B: formats whose leaf edit costs grow with the document per operation
#      (DOC, ODG, ODB, ODF formula): fixture of 1,000 targets, 100 / 1,000.
#   C: ODI (one frame): its 6 independent frame properties.
#   D: XLS with a fixture of exactly N cells (record 0209's shape: N
#      operations on an N-cell column), eager and file-backed owners.
set -u
BIN=$1 OUT=$2 CPU=${CPU:-30} TAG=${TAG:-run}
mkdir -p "$OUT"
log() { echo "== $* $(date +%T) load $(cut -d' ' -f1 /proc/loadavg)" >> "$OUT/log.txt"; }
run() { local name=$1; shift; log "$name $*"; taskset -c "$CPU" "$BIN" "$@" >> "$OUT/$name" 2>> "$OUT/log.txt" || echo "FAIL $name" >> "$OUT/log.txt"; }
SETS=${SETS:-ABCD}
A=${FORMATS_A:-docx,xlsx,pptx,xlsb,xls,ppt,rtf,odt,oth,odm,ods,odp,odc,pages,keynote,numbers}
for trial in 1 2 3; do
  [[ $SETS == *A* ]] && run batch-A-$TAG-$trial.json --formats $A --counts 100,1000,10000 --rounds 5
  [[ $SETS == *B* ]] && run batch-B-$TAG-$trial.json --formats doc,odg,odb,odf --counts 100,1000 --rounds 5
  [[ $SETS == *C* ]] && run batch-C-$TAG-$trial.json --formats odi --rounds 5
  [[ $SETS == *D* ]] && for n in 100 1000 10000; do
    run batch-D-$TAG-$trial-n$n.json --formats xls --counts $n --rounds 5
  done
done
echo done >> "$OUT/log.txt"
