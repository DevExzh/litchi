#!/bin/bash
# Change 0754 gate runner: run one gate command in the branch worktree, append
# the command, exit status, duration and a summary tail to gates-raw.txt.
set -u
S=/home/zhuhe/code/litchi-worktrees/scratch/0754
cd /home/zhuhe/code/litchi-worktrees/0754-docx-semantic-edit-and-text-path
export TMPDIR=$S/tmp CARGO_BUILD_JOBS=6 CARGO_TARGET_DIR=/home/zhuhe/code/litchi-worktrees/targets/0754
label=$1; shift
log=$S/gates/$label.log
start=$(date +%s)
"$@" > "$log" 2>&1
status=$?
{
  echo "## $label"
  echo "command: $*"
  echo "exit: $status  ($(( $(date +%s) - start ))s)"
  if grep -q "^test result" "$log"; then
    grep -E "^test result" "$log" | awk '{p+=$4; f+=$6; i+=$8} END {print "tests: passed",p,"failed",f,"ignored",i}'
  fi
  grep -E "^(warning|error)(\[|:)" "$log" | sort | uniq -c | head -5
  tail -2 "$log"
  echo
} >> $S/gates/gates-raw.txt
echo "$label exit $status"
