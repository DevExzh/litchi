#!/bin/bash
# Differential: the full-text projection and the whole-sheet walk of every XLS
# fixture, as an outcome digest plus the counted observations.
BIN="$1"; ROOT="$2"; OUT="$3"
: > "$OUT"
for f in "$ROOT"/test-data/ole/xls/*.xls "$ROOT"/test-data/poi/test-data/spreadsheet/*.xls; do
  [ -f "$f" ] || continue
  name="${f#"$ROOT/"}"
  for op in full-text all-cells; do
    json=$(taskset -c 16 "$BIN" --input "$f" --mode file-source --operation "$op" \
      --warmups 1 --samples 1 --worksheet-index 0 2>&1)
    echo "$json" | python3 -c "
import json,sys
raw=sys.stdin.read()
try:
    d=json.loads(raw)
    r=d['records'][-1]
    print('$name','$op','outcome='+str(r['observation'].get('outcome')),'reads=%d'%r['metrics']['read_calls'],'bytes=%d'%r['metrics']['read_bytes'],'versions=%d'%r['metrics']['version_calls'])
except Exception:
    print('$name','$op','tool_error='+raw.strip().replace('\n',' ')[:100])
" >> "$OUT"
  done
done
