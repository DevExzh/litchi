#!/bin/sh
# abba.sh BIN_DIR FIXTURE OUT_TSV [ROUNDS] : binary-level ABBA of the
# scoped_patch_probe (record 0357a); one line per run with peak RSS.
bin=$1; fixture=$2; out=$3; rounds=${4:-10}
printf 'round\torder\tbuild\tmax_rss_kb\texit\tjson\n' > "$out"
r=1
while [ $r -le $rounds ]; do
  o=1
  for b in before after after before; do
    tmp=$(mktemp)
    json=$(/usr/bin/time -f '%M' -o "$tmp" taskset -c 29 "$bin/$b" "$fixture" 1000)
    code=$?
    printf '%s\t%s\t%s\t%s\t%s\t%s\n' $r $o $b "$(cat "$tmp")" $code "$json" >> "$out"
    rm -f "$tmp"
    o=$((o+1))
  done
  r=$((r+1))
done
