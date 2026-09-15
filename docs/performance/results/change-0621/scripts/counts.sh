#!/bin/bash
# $1 = binary, $2 = repo root (for fixtures), $3 = output dir, $4 = leg label
BIN="$1"; ROOT="$2"; OUT="$3"; LEG="$4"
mkdir -p "$OUT"
run() { # mode operation fixture label extra...
  local m="$1" op="$2" f="$3" lab="$4"; shift 4
  local j="$OUT/$LEG-$lab-$m-$op.json"
  taskset -c 16 "$BIN" --input "$ROOT/$f" --mode "$m" --operation "$op" --warmups 1 --samples 1 "$@" > "$j" 2> "$j.err"
  python3 - "$j" "$m" "$op" "$lab" <<'PY'
import json,sys
p,m,op,lab=sys.argv[1:5]
try:
    d=json.load(open(p))
except Exception:
    print(f"{lab:6s} {m:12s} {op:11s} ERROR {open(p+'.err').read().strip()[:80]}")
    sys.exit()
r=d['records'][-1]; g=r['metrics']; o=r['observation']
out=o.get('outcome')
print(f"{lab:6s} {m:12s} {op:11s} version={g['version_calls']:6d} len={g['len_calls']:3d} reads={g['read_calls']:5d} bytes={g['read_bytes']:8d} outcome={out}")
PY
}
for m in file-source facade-file owned-readat; do
  run "$m" open   test-data/poi/test-data/spreadsheet/54016.xls 54016
  run "$m" list   test-data/poi/test-data/spreadsheet/54016.xls 54016
  run "$m" one-cell test-data/poi/test-data/spreadsheet/54016.xls 54016 --worksheet-index 0 --row 1 --column 0
done
for m in file-source owned-readat; do
  run "$m" all-cells test-data/poi/test-data/spreadsheet/54016.xls 54016 --worksheet-index 0 --all-cells-strategy scan
  run "$m" full-text test-data/poi/test-data/spreadsheet/54016.xls 54016
done
for m in file-source facade-file; do
  run "$m" open   test-data/ole/xls/ConditionalFormattingSamples.xls cfs
  run "$m" one-cell test-data/ole/xls/ConditionalFormattingSamples.xls cfs --worksheet-index 0 --row 1 --column 0
done
run file-source full-text test-data/ole/xls/ConditionalFormattingSamples.xls cfs
run file-source full-text test-data/ole/xls/SimpleMultiCell.xls simple
run file-source open test-data/ole/xls/SimpleMultiCell.xls simple
