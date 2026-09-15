#!/bin/bash
BIN="$1"; ROOT="$2"; OUT="$3"; LEG="$4"
mkdir -p "$OUT"
pair() { # mode op fixture label
  local m="$1" op="$2" f="$3" lab="$4"
  for n in 1 6; do
    taskset -c 16 valgrind --tool=callgrind --callgrind-out-file="$OUT/$LEG-$lab-$m-$op-s$n.out" \
      --cache-sim=no --branch-sim=no \
      "$BIN" --input "$ROOT/$f" --mode "$m" --operation "$op" --warmups 1 --samples $n \
      > /dev/null 2> "$OUT/$LEG-$lab-$m-$op-s$n.log"
  done
  python3 - "$OUT/$LEG-$lab-$m-$op-s1.log" "$OUT/$LEG-$lab-$m-$op-s6.log" "$m" "$op" "$lab" "$LEG" <<'PY'
import sys,re
def ir(path):
    for line in open(path):
        m=re.search(r"Collected\s*:\s*([0-9,]+)", line)
        if m: return int(m.group(1).replace(",",""))
        m=re.search(r"refs:\s*([0-9,]+)", line)
        if m: return int(m.group(1).replace(",",""))
    raise SystemExit("no Ir in "+path)
a,b=ir(sys.argv[1]),ir(sys.argv[2])
m,op,lab,leg=sys.argv[3:7]
print(f"{leg:6s} {lab:6s} {m:12s} {op:10s} Ir_per_sample={(b-a)/5:,.0f}  (s1={a:,} s6={b:,})")
PY
}
pair file-source open      test-data/poi/test-data/spreadsheet/54016.xls 54016
pair file-source full-text test-data/poi/test-data/spreadsheet/54016.xls 54016
pair facade-file open      test-data/poi/test-data/spreadsheet/54016.xls 54016
