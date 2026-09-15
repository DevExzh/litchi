#!/usr/bin/env bash
SC=/tmp/claude-1001/-home-zhuhe-code-litchi/709d31e6-bf68-4968-8b5c-9a2af1c22cc8/scratchpad/agents/0589
cd /home/zhuhe/code/litchi-worktrees/0589
run() { name=$1; shift; echo "### $* ###" > $SC/gates/$name.txt; nice -n 5 "$@" >> $SC/gates/$name.txt 2>&1; echo "EXIT=$?" >> $SC/gates/$name.txt; }
run fmt cargo fmt --all --check
run clippy cargo clippy -p litchi-cfb -p litchi-doc -p litchi-ppt --all-targets
run clippy-consumers cargo clippy -p litchi-xls -p litchi-ole-common --all-targets
run test cargo test -p litchi-cfb -p litchi-doc -p litchi-ppt
run test-consumers cargo test -p litchi-xls -p litchi-ole-common
run doc cargo doc -p litchi-cfb -p litchi-doc -p litchi-ppt --no-deps
echo GATESDONE
