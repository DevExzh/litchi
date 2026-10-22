#!/usr/bin/env bash
# Callgrind Ir for the protected matched-v5 constructor binaries; per-op = (Ir(12000)-Ir(2000))/10000.
set -euo pipefail
V5=/home/zhuhe/code/litchi-worktrees/.coord/constructor-0361f-coordinator/matched-v5
cd "$(dirname "$0")"
sha256sum $V5/build-base/constructor-native $V5/build-candidate/constructor-native > binaries.sha256
for route in field_path field_info message_info archive_info archive_object archive_object_reject parse validate encode \
    retained_parse retained_validate retained_encoded_len retained_encode retained_validate_tight retained_encode_tight; do
  for side in base candidate; do
    for n in 2000 12000; do
      out=$route-$side-$n
      valgrind --tool=callgrind --callgrind-out-file=$out.out $V5/build-$side/constructor-native $route $n > $out.stdout 2> $out.stderr
      echo "$route $side $n $(grep -m1 '^summary:' $out.out | awk '{print $2}') $(python3 -c 'import json,sys; d=json.load(open(sys.argv[1])); print(d["output_sha256"], d["checksum"])' $out.stdout)"
    done
  done
done
