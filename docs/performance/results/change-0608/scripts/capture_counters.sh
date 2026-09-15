#!/usr/bin/env bash
# capture_counters.sh <out-file>
#
# The deterministic logical counters, taken first: reads, read bytes, source
# observations, len calls and seeks for every leg, fixture, in-memory mode and
# operation.  This is the control for change 0608: if the two measurement
# scaffolds read exactly what the base reads, then the SST bytes are read at
# open whatever the index does with them, and deferral is a CPU question, not
# an I/O one.
set -euo pipefail
SC=${SC:-/tmp/claude-1001/-home-zhuhe-code-litchi/709d31e6-bf68-4968-8b5c-9a2af1c22cc8/scratchpad/agents/0608}
REPO=${REPO:-/home/zhuhe/code/litchi}
OUT=$1
: > "$OUT"
cells=(
  "flagship:$REPO/test-data/ole/xls/ConditionalFormattingSamples.xls:1"
  "cv:$REPO/test-data/ole/xls/WithCustomViews.xls:1"
  "54016:$REPO/test-data/poi/test-data/spreadsheet/54016.xls:0"
)
for leg in base nostore nowalk; do
  for cell in "${cells[@]}"; do
    IFS=: read -r stem path sheet <<<"$cell"
    for mode in owned-readat file-source; do
      for op in open list one-cell; do
        "$SC/bin/xsa-$leg" --input "$path" --mode "$mode" --operation "$op" \
          --worksheet-index "$sheet" --warmups 1 --samples 3 2>/dev/null \
        | python3 -c "
import json,sys
d=json.load(sys.stdin)
ms=[r['metrics'] for r in d['records']]
keys=('read_calls','read_bytes','version_calls','len_calls','seek_calls')
uniq={tuple(m[k] for k in keys) for m in ms}
proj=d['semantic_oracle']['source_implementation_projection']
print('$leg','$stem','$mode','$op', 'IDENTICAL' if len(uniq)==1 else 'VARIES',
      dict(zip(keys,sorted(uniq)[0])), 'sheets=%d' % proj['worksheet_count'],
      'cell=%s' % json.dumps(proj['selected_cell']))
" >> "$OUT"
      done
    done
  done
done
