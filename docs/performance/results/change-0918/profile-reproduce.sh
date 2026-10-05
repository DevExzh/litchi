#!/bin/sh
set -eu

binary=/home/zhuhe/litchi-performance-0917/target/release/doc_retention_probe
floating=/home/zhuhe/code/litchi/test-data/ole/doc/FloatingPictures.doc
no_head_foot=/home/zhuhe/code/litchi/test-data/ole/doc/NoHeadFoot.doc
profile_root=${LITCHI_PROFILE_0918_ROOT:-/home/zhuhe/litchi-performance-0918/reproduce}

if [ -e "$profile_root" ]; then
    printf '%s\n' "refusing to overwrite existing output directory: $profile_root" >&2
    exit 2
fi
mkdir "$profile_root"

test "$(sha256sum "$binary" | awk '{print $1}')" = \
  2ae48b4f7548c1468c8e7691d56d89b231c4a63f52ed48383f07fe3609c59a13
test "$(sha256sum "$floating" | awk '{print $1}')" = \
  41cc0cd8f2d7390266f844f5bddbcd29d08ac5338abc33832d9c9a6ed5d0f85d
test "$(sha256sum "$no_head_foot" | awk '{print $1}')" = \
  45e5df073f34314da6f39d2dad119fb2ef23470878fd2df67f632864cd92ea48

run_profile() {
    name=$1
    input=$2
    case_name=$3
    edits=$4
    run_dir=$profile_root/$name
    mkdir "$run_dir"
    date -u +%Y-%m-%dT%H:%M:%S.%NZ > "$run_dir/start.utc"
    date +%s%N > "$run_dir/start.epoch-ns"
    set +e
    /usr/bin/taskset -c 14 valgrind --tool=callgrind \
        --instr-atstart=yes --collect-atstart=no \
        --toggle-collect='*measure_one*' --collect-jumps=yes \
        --simulate-cache=yes --dump-instr=yes \
        --callgrind-out-file="$run_dir/callgrind.out" \
        "$binary" --case "$case_name" --input "$input" \
        --edits "$edits" --retention default --warmup 0 --samples 1 \
        > "$run_dir/probe.stdout" 2> "$run_dir/probe.stderr"
    status=$?
    set -e
    printf '%s\n' "$status" > "$run_dir/exit"
    date -u +%Y-%m-%dT%H:%M:%S.%NZ > "$run_dir/finish.utc"
    date +%s%N > "$run_dir/finish.epoch-ns"
    test "$status" -eq 0
}

run_profile callgrind-docfloat-e02-default "$floating" docfloat 2
run_profile callgrind-docfloat-e32-default "$floating" docfloat 32
run_profile callgrind-docnohf-e32-default "$no_head_foot" docnohf 32
