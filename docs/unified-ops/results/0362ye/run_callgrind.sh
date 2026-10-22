#!/bin/sh
set -eu
bin=$1
mode=$2
outdir=$3
mkdir -p "$outdir"
for route in open first_slide repeat_slide first_text repeat_text write_text repeat_write_text unchanged_write append_commit_write replay; do
  out="$outdir/$mode-$route.callgrind"
  log="$outdir/$mode-$route.stdout"
  err="$outdir/$mode-$route.stderr"
  valgrind --tool=callgrind --quiet --callgrind-out-file="$out" "$bin" "$route" >"$log" 2>"$err"
  collected=$(awk '$1 == "summary:" {print $2; exit}' "$out")
  printf '%s %s %s\n' "$mode" "$route" "$collected"
done
