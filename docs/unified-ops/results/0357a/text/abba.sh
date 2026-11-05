#!/bin/sh
# abba.sh BIN_DIR DECK OUT_TSV [ROUNDS] [ITERATIONS] : binary-level ABBA of
# text_splice_probe (record 0357a). Every slot runs mode `primary`; the
# `after` slots also run mode `text`, which the before build lacks.
bin=$1; deck=$2; out=$3; rounds=${4:-10}; iterations=${5:-50}
printf 'round\torder\tbuild\tmode\tmax_rss_kb\texit\tjson\n' > "$out"
r=1
while [ $r -le $rounds ]; do
  o=1
  for b in before after after before; do
    modes=primary
    [ $b = after ] && modes="primary text"
    for m in $modes; do
      tmp=$(mktemp)
      json=$(/usr/bin/time -f '%M' -o "$tmp" taskset -c 29 "$bin/$b" $m "$deck" $iterations)
      code=$?
      printf '%s\t%s\t%s\t%s\t%s\t%s\t%s\n' $r $o $b $m "$(cat "$tmp")" $code "$json" >> "$out"
      rm -f "$tmp"
    done
    o=$((o+1))
  done
  r=$((r+1))
done
