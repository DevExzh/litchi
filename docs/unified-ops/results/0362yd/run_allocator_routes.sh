#!/bin/sh
set -eu
bin=$1
mode=$2
outdir=$3
mkdir -p "$outdir"
for route in open first_slide repeat_slide first_text repeat_text write_text repeat_write_text unchanged_write append_commit_write replay; do
  "$bin" "$route" > "$outdir/$mode-$route.json"
done
