#!/bin/sh
# range_abba.sh BIN_DIR DECK OUT_TSV [ROUNDS] : binary-level ABBA of the
# 0627 ole2_range_source PPT cases of tools/perf-baseline (record 0357a).
# One line per case and run: p50/p95 elapsed ns, the first sample's reads and bytes.
bin=$1; deck=$2; out=$3; rounds=${4:-5}
cases=ppt_range_source_open,ppt_range_source_open_one_shape_text,ppt_owned_source_control_open,ppt_owned_source_control_open_one_shape_text
printf 'round\torder\tbuild\tcase\tp50_ns\tp95_ns\tread_calls\tread_bytes\n' > "$out"
r=1
while [ $r -le $rounds ]; do
  o=1
  for b in before after after before; do
    json=$(mktemp)
    taskset -c 29 "$bin/perf-$b" --warmup 20 --samples 50 --case $cases \
      --ole2-file "$deck" --range-fixed-latency-us 1000 --range-request-overhead-us 0 \
      --range-bandwidth-bytes-per-sec 104857600 --range-max-physical-bytes 65536 \
      --json "$json" > /dev/null 2>&1 || echo "run $r $b failed" >&2
    python3 -I -c '
import json, sys
d = json.load(open(sys.argv[1]))
for res in d["results"]:
    e, s = res["elapsed_ns"], res.get("source") or {}
    first = lambda v: v[0] if isinstance(v, list) and v else v
    print("\t".join(map(str, [sys.argv[2], sys.argv[3], sys.argv[4], res["case"], e["p50"], e["p95"], first(s.get("read_calls")), first(s.get("read_bytes"))])))
' "$json" $r $o $b >> "$out"
    rm -f "$json"
    o=$((o+1))
  done
  r=$((r+1))
done
