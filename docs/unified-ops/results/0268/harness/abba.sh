#!/bin/bash
set -u
S=${BENCH:?}
T=${TARGET:?}
O=${O:-$S/out}; R=${R:-$S/runs}; mkdir -p $O $R; rm -f $R/*
CPU=${CPU:-7}
i=0
for order in A B B A A B B A A B B A; do
  i=$((i+1))
  if [ $order = A ]; then bin=$T/perf-before/release/odf-w3-bench-before; tag=before; else bin=$T/perf-after/release/odf-w3-bench-after; tag=after; fi
  taskset -c $CPU $bin 25 $O > $R/$(printf %02d $i)-$tag.tsv 2>/dev/null || { echo "run $i failed"; exit 1; }
done
echo done
