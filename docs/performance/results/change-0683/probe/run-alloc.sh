#!/bin/sh
set -eu

if [ "$#" -lt 4 ]; then
    echo "usage: run-alloc.sh BEFORE_ALLOC AFTER_ALLOC CORPUS_DIR OUT_DIR [SHEET]" >&2
    exit 2
fi

before=$1
after=$2
corpus_dir=$3
out_dir=$4
sheet=${5:-Sheet1}
mkdir -p "$out_dir"

cases='dense sparse long-inline shared-string formula late-fallback late-refusal escaped-text malformed-escape'
operations='visit-cold cells-cold visit-selected cells-selected visit-warm cells-warm'

run_leg() {
    leg=$1
    binary=$2
    for case in $cases; do
        input=$corpus_dir/$case.xlsx
        if [ ! -f "$input" ]; then
            echo "missing corpus: $input" >&2
            exit 1
        fi
        for operation in $operations; do
            case "$case:$operation" in late-refusal:*-warm|malformed-escape:*-warm) continue ;; esac
            "$binary" "$operation" "$input" "$sheet" \
                > "$out_dir/${leg}-${case}-${operation}.tsv"
        done
    done
}

run_leg before "$before"
run_leg after "$after"
