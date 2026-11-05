#!/usr/bin/env bash
# usage: cg.sh OUTDIR BIN_LABEL... ; env ROUTES, COUNTS (message counts), default 8
# Ir/op = (Ir(12000) - Ir(2000)) / 10000 ; runs sequentially.
set -euo pipefail
M=/deleted-coord-artifact
out=$1; shift
mkdir -p "$out"
ROUTES=${ROUTES:-"field_path field_info message_info archive_info archive_info_raw archive_object archive_object_reject parse validate encode retained_parse retained_validate retained_encoded_len retained_encode retained_validate_tight retained_encode_tight archive_info_e2e archive_info_generic_e2e archive_info_raw_e2e archive_object_serialize_e2e"}
COUNTS=${COUNTS:-8}
for route in $ROUTES; do for c in $COUNTS; do for side in "$@"; do
  r=()
  for n in 2000 12000; do
    f=$out/$route-$c-$side-$n
    IWA_MESSAGE_COUNT=$c valgrind --tool=callgrind --callgrind-out-file=$f.out $M/bin/$side $route $n > $f.json 2> $f.stderr
    r+=($(grep -m1 '^summary:' $f.out | awk '{print $2}'))
  done
  sha=$(python3 -c 'import json,sys; d=json.load(open(sys.argv[1])); print(d["output_sha256"][:12], d["checksum"])' $out/$route-$c-$side-12000.json)
  python3 -c "import sys; a,b=int(sys.argv[1]),int(sys.argv[2]); print(f'$route\t$c\t$side\t{a}\t{b}\t{(b-a)/10000:.4f}\t$sha')" ${r[0]} ${r[1]}
done; done; done | tee -a $out/summary.tsv
