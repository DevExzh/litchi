#!/usr/bin/env bash
# Run one gate in a directory with a target dir and append the command, exit
# code and output tail to the gates log. Usage: gate.sh <log> <dir> <target> <cmd...>
set -u
OUT=$1; DIR=$2; TARGET=$3; shift 3
export CARGO_BUILD_JOBS=6 TMPDIR=/home/zhuhe/code/litchi-worktrees/scratch/0743/tmp
LOG=$(mktemp -p "$TMPDIR" gate.XXXXXX)
(cd "$DIR" && CARGO_TARGET_DIR=$TARGET "$@") > "$LOG" 2>&1
CODE=$?
{
  echo "### $(cd "$DIR" && git rev-parse --short HEAD) :: $*"
  echo "cwd: $DIR"
  echo "exit: $CODE"
  grep -E "^(error|warning: unused|Summary|OK|PASS|FAIL)" "$LOG" | tail -20
  if grep -q "^test result" "$LOG"; then
    grep -c "^test result" "$LOG" | sed 's/^/test result lines: /'
    awk '/^test result/ {p += $4; f += $6; i += $8} END {print "totals: passed=" p " failed=" f " ignored=" i}' "$LOG"
  fi
  tail -4 "$LOG"
  echo
} >> "$OUT"
rm -f "$LOG"
echo "exit $CODE: $*"
