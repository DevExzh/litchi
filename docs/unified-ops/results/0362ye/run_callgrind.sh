#!/bin/sh
set -eu
bin=$1
mode=$2
outdir=$3
version=${4:-after}
case "$version" in
  before) expected_mode="0358a-pre-change-e4a0614f5d" ;;
  after) expected_mode="UNBUILT-0362YE" ;;
  *) echo "version must be before or after" >&2; exit 2 ;;
esac
mkdir -p "$outdir"
for route in open first_slide repeat_slide first_text repeat_text write_text repeat_write_text unchanged_write append_commit_write replay; do
  out="$outdir/$mode-$route.callgrind"
  log="$outdir/$mode-$route.stdout"
  err="$outdir/$mode-$route.stderr"
  valgrind --tool=callgrind --quiet --callgrind-out-file="$out" "$bin" "$route" >"$log" 2>"$err"
  actual_mode=$(sed -n 's/.*"mode":"\([^"]*\)".*/\1/p' "$log")
  test "$actual_mode" = "$expected_mode" || {
    echo "probe mode mismatch: expected $expected_mode, got $actual_mode" >&2
    exit 1
  }
  collected=$(awk '$1 == "summary:" {print $2; exit}' "$out")
  printf '%s %s %s\n' "$mode" "$route" "$collected"
done
