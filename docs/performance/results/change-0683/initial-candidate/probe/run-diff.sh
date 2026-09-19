#!/bin/sh
set -eu

if [ "$#" -lt 4 ]; then
    echo "usage: run-diff.sh BEFORE AFTER CORPUS_DIR OUT_DIR [SHEET]" >&2
    exit 2
fi

before=$1
after=$2
corpus_dir=$3
out_dir=$4
sheet=${5:-Sheet1}
mkdir -p "$out_dir"

cases='dense sparse long-inline shared-string formula late-fallback late-refusal'
for case in $cases; do
    input=$corpus_dir/$case.xlsx
    "$before" diff "$input" "$sheet" > "$out_dir/before-${case}.tsv"
    "$after" diff "$input" "$sheet" > "$out_dir/after-${case}.tsv"
    diff -u "$out_dir/before-${case}.tsv" "$out_dir/after-${case}.tsv"
done
