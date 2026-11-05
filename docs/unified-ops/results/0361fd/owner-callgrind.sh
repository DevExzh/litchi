#!/usr/bin/env bash
# callgrind.sh : total Ir for each route/endpoint at N=1 and N=5; per-op = delta/4.
set -euo pipefail
cd "$(dirname "$0")/.."
for route in numbers_headers keynote_media pages_footnote_insert pages_section_name; do
  for side in base candidate; do
    for n in 1 5; do
      out=callgrind/$route-$side-$n
      systemd-run --user --scope --quiet --collect -p OOMPolicy=continue -p MemoryMax=20G -p MemorySwapMax=0 -- \
        valgrind --tool=callgrind --callgrind-out-file=$out.out build-$side/owner-probe $route $n fixtures \
        > $out.stdout 2> $out.stderr
      ir=$(grep -m1 '^summary:' $out.out | awk '{print $2}')
      echo "$route $side $n $ir $(cat $out.stdout)"
    done
  done
done
