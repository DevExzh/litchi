#!/bin/bash
# Run the remaining gates and print exact counts.
set -u
cd /home/zhuhe/code/litchi/.claude/worktrees/agent-a4d3a0dc180312e39 || exit 1
S=/tmp/claude-1001/-home-zhuhe-code-litchi/4b20edf7-6b1c-40a9-9e79-3ca6a15219f0/scratchpad/prepared

echo "########## GATE 2: leaf crates ##########"
cargo test --offline --locked -p litchi-opc -p litchi-docx -p litchi-pptx -p litchi-xlsx \
  --all-features -- --test-threads=2 >"$S/gate2.txt" 2>&1
echo "exit=$?"
grep -E "^test result:" "$S/gate2.txt" | awk '{p+=$4; f+=$6; i+=$8} END {print "  passed="p" failed="f" ignored="i}'
grep -E "^(error|failures:)" "$S/gate2.txt" | head -10

echo "########## GATE 3: -p litchi ##########"
cargo test --offline --locked -p litchi --all-features --no-fail-fast -- --test-threads=2 \
  >"$S/gate3.txt" 2>&1
echo "exit=$?"
grep -E "^test result:" "$S/gate3.txt" | awk '{p+=$4; f+=$6; i+=$8} END {print "  passed="p" failed="f" ignored="i}'
echo "  --- failing test names ---"
awk '/^failures:$/{flag=1;next} /^test result:/{flag=0} flag && /^    /{print "   "$0}' "$S/gate3.txt" | sort -u

echo "########## GATE 4: clippy ##########"
cargo clippy --offline --locked -p litchi-opc -p litchi --all-targets --all-features -- -D warnings \
  >"$S/gate4.txt" 2>&1
echo "exit=$?"
grep -cE "^(error|warning)" "$S/gate4.txt" | sed 's/^/  error+warning lines: /'
grep -E "^(error|warning)" "$S/gate4.txt" | head -20

echo "########## GATE 5: rustdoc ##########"
RUSTDOCFLAGS="-D warnings" cargo doc --offline --locked -p litchi-opc -p litchi \
  --all-features --no-deps >"$S/gate5.txt" 2>&1
echo "exit=$?"
grep -E "^(error|warning)" "$S/gate5.txt" | head -20

echo "########## DONE ##########"
