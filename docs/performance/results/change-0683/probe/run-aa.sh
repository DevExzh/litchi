#!/bin/sh
set -eu

if [ "$#" -lt 5 ]; then
    echo "usage: run-aa.sh BINARY CORPUS_DIR OUT_DIR WARMUP SAMPLES [SHEET]" >&2
    exit 2
fi

binary=$1
corpus_dir=$2
out_dir=$3
warmup=$4
samples=$5
sheet=${6:-Sheet1}
mkdir -p "$out_dir"

cases='dense sparse long-inline shared-string formula late-fallback late-refusal escaped-text malformed-escape'
operations='visit-cold cells-cold visit-selected cells-selected visit-warm cells-warm'

run_leg() {
    label=$1
    for case in $cases; do
        input=$corpus_dir/$case.xlsx
        for operation in $operations; do
            case "$case:$operation" in late-refusal:*-warm|malformed-escape:*-warm) continue ;; esac
            "$binary" bench "$operation" "$input" "$sheet" "$warmup" "$samples" \
                > "$out_dir/${label}-${case}-${operation}.tsv"
        done
    done
}

run_leg a1
run_leg a2
