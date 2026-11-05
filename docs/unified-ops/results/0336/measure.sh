#!/usr/bin/env bash
# Record 0336: File::dump against a read-only node-tree traversal.
# Build at the record's commit, then run three pinned processes.
set -euo pipefail
cargo build --release -p litchi-ops --example dump_bench --features docx,xlsx,pptx
sha256sum target/release/examples/dump_bench
for i in 1 2 3; do taskset -c 7 target/release/examples/dump_bench; done | tee dump_bench.jsonl
# The corpus survey (release; files up to 512 KiB):
# DUMP_SURVEY=docx,xlsx,xlsb,pptx,odt,ods,rtf,doc,xls \
#   cargo test --release -p litchi-ops --all-features --test dump_round_trip survey -- --ignored --nocapture
