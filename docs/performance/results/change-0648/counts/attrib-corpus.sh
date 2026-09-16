#!/bin/bash
# Independent corpus differential through change 0605's xls_source_attribution:
# every fixture, two operations, one sample, outcome and counters per row.
S=/tmp/claude-1001/-home-zhuhe-code-litchi/709d31e6-bf68-4968-8b5c-9a2af1c22cc8/scratchpad/agents/0648
leg=$1; tree=$2; out=$3
cd "$tree" || exit 1
: > "$out"
while read -r path; do
  for op in all-cells full-text; do
    line=$(taskset -c 23 "$S/bin/attrib-$leg" --input "$path" --mode owned-readat \
      --operation "$op" --worksheet-index 0 --all-cells-strategy scan \
      --warmups 1 --samples 1 2>/dev/null | python3 -c "
import sys, json
try:
    d = json.load(sys.stdin)
except Exception:
    print('no-report'); raise SystemExit
r = d['records'][-1]
m = r['metrics']
print('reads=%d bytes=%d versions=%d outcome=%s' % (m['read_calls'], m['read_bytes'], m['version_calls'], r['observation']['outcome']))
")
    [ -z "$line" ] && line="refused-or-empty"
    printf '%s\t%s\t%s\n' "$path" "$op" "$line" >> "$out"
  done
done < "$S/counts/xls-fixtures.txt"
