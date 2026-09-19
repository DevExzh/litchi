#!/bin/sh
set -eu

if [ "$#" -lt 6 ]; then
    echo "usage: run-timing.sh BEFORE AFTER CORPUS_DIR OUT_DIR WARMUP SAMPLES [SHEET]" >&2
    exit 2
fi

before=$1
after=$2
corpus_dir=$3
out_dir=$4
warmup=$5
samples=$6
sheet=${7:-Sheet1}
mkdir -p "$out_dir"

cases='dense sparse long-inline shared-string formula late-fallback late-refusal escaped-text malformed-escape'
operations='visit-cold cells-cold visit-selected cells-selected visit-warm cells-warm'

run_leg() {
    label=$1
    binary=$2
    for case in $cases; do
        input=$corpus_dir/$case.xlsx
        for operation in $operations; do
            case "$case:$operation" in late-refusal:*-warm|malformed-escape:*-warm) continue ;; esac
            "$binary" bench "$operation" "$input" "$sheet" "$warmup" "$samples" \
                > "$out_dir/${label}-${case}-${operation}.tsv"
        done
    done
}

# The two A legs bracket the two B legs. The output names preserve the pair
# identity so a later summarizer can reject mixed or missing samples.
run_leg a1 "$before"
run_leg b1 "$after"
run_leg b2 "$after"
run_leg a2 "$before"
