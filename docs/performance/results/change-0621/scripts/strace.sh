#!/bin/bash
# $1 binary, $2 repo root, $3 outdir, $4 leg
BIN="$1"; ROOT="$2"; OUT="$3"; LEG="$4"
mkdir -p "$OUT"
run() { # mode op fixture label
  local m="$1" op="$2" f="$3" lab="$4"; shift 4
  for n in 1 11; do
    taskset -c 16 strace -f -c -e trace=statx,pread64,newfstatat,openat \
      -o "$OUT/$LEG-$lab-$m-$op-s$n.txt" \
      "$BIN" --input "$ROOT/$f" --mode "$m" --operation "$op" --warmups 1 --samples $n "$@" > /dev/null 2>&1
  done
  python3 - "$OUT/$LEG-$lab-$m-$op-s1.txt" "$OUT/$LEG-$lab-$m-$op-s11.txt" "$m" "$op" "$lab" <<'PY'
import sys, re
def counts(path):
    out={}
    for line in open(path):
        parts=line.split()
        if len(parts)>=4 and parts[-1] in ("statx","pread64","newfstatat","openat"):
            out[parts[-1]]=int(parts[-2]) if parts[-2].isdigit() else int(parts[-3])
    return out
a,b,m,op,lab=sys.argv[1],sys.argv[2],sys.argv[3],sys.argv[4],sys.argv[5]
ca,cb=counts(a),counts(b)
row=[]
for call in ("statx","newfstatat","pread64","openat"):
    x,y=ca.get(call,0),cb.get(call,0)
    row.append(f"{call}={(y-x)/10:.1f}")
print(f"{lab:6s} {m:12s} {op:10s} per-sample " + " ".join(row))
PY
}
run facade-file open  test-data/poi/test-data/spreadsheet/54016.xls 54016
run file-source open  test-data/poi/test-data/spreadsheet/54016.xls 54016
run file-source full-text test-data/poi/test-data/spreadsheet/54016.xls 54016
run facade-file open  test-data/ole/xls/ConditionalFormattingSamples.xls cfs
run file-source open  test-data/ole/xls/ConditionalFormattingSamples.xls cfs
