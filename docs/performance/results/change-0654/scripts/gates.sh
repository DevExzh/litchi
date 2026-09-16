#!/bin/bash
S=/tmp/claude-1001/-home-zhuhe-code-litchi/709d31e6-bf68-4968-8b5c-9a2af1c22cc8/scratchpad/agents/0654/gates
cd /home/zhuhe/code/litchi-worktrees/0654
export CARGO_BUILD_JOBS=12
run() { name=$1; shift; echo "### $name"; nice -n 10 "$@" > $S/$name.txt 2>&1; echo "$name exit=$?"; }
run fmt            cargo fmt --all --check
run clippy         cargo clippy -p litchi-opc -p xml-minifier --all-targets
run test-opc       cargo test -p litchi-opc -p xml-minifier
run doc            cargo doc -p litchi-opc -p xml-minifier --no-deps
run test-consumers cargo test -p litchi-xlsx -p litchi-docx -p litchi-pptx
run test-facade    cargo test -p litchi --features docx,xlsx,pptx,xls
run non-iwork      python3 tools/non_iwork_gate.py verify
