#!/bin/bash
# Unit-level probe: the base-code and candidate probe binaries on the exact
# bytes captured from the measured publication, A B B A per round, CPU 24.
set -u
S=/home/zhuhe/code/litchi-worktrees/scratch/0747
G=$S/gdbcap
OUT=$S/probe/abba
mkdir -p "$OUT"
args() {
  echo "$1-worksheet $G/$1/hit046.bin $G/$1/hit047.bin $1-workbook $G/$1/hit044.bin $G/$1/hit045.bin"
}
for round in 1 2 3 4; do
  for shape in medium dense-sparse; do
    # shellcheck disable=SC2046
    taskset -c 24 $S/bin/audit_probe_0747.before $(args $shape) > "$OUT/$shape-r$round-s1-A.csv"
    taskset -c 24 $S/bin/audit_probe_0747.after $(args $shape) > "$OUT/$shape-r$round-s2-B.csv"
    taskset -c 24 $S/bin/audit_probe_0747.after $(args $shape) > "$OUT/$shape-r$round-s3-B.csv"
    taskset -c 24 $S/bin/audit_probe_0747.before $(args $shape) > "$OUT/$shape-r$round-s4-A.csv"
  done
done
echo done > "$OUT/status.txt"
