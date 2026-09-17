#!/usr/bin/env bash
# Change 0673's open-differential oracle.
#
# Usage: run.sh [before-tree] [after-tree] [file-list]
set -euo pipefail

HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
BEFORE_TREE="${1:-/home/zhuhe/code/litchi-worktrees/before-5fa92d7ce}"
AFTER_TREE="${2:-/home/zhuhe/code/litchi-worktrees/0673}"
LIST="${3:-$HERE/../../change-0632/census/zip-containers.txt}"
WORK="$(mktemp -d /tmp/litchi-0673-oracle.XXXXXX)"

for leg in before after; do
  case "$leg" in
    before) TREE="$BEFORE_TREE" ;;
    after) TREE="$AFTER_TREE" ;;
  esac
  mkdir -p "$WORK/oracle-$leg/src"
  cp "$HERE/src/main.rs" "$WORK/oracle-$leg/src/main.rs"
  sed "s#@TREE@#$TREE#g" "$HERE/Cargo.toml.template" > "$WORK/oracle-$leg/Cargo.toml"
  echo "== building $leg oracle against $TREE"
  (
    cd "$WORK/oracle-$leg"
    CARGO_PROFILE_RELEASE_DEBUG=0 CARGO_NET_OFFLINE=true \
      CARGO_TARGET_DIR="$WORK/target-$leg" cargo build --release -j 2
  )
  echo "== running $leg"
  (
    cd "$AFTER_TREE"
    taskset -c 8 "$WORK/target-$leg/release/opc-open-oracle" \
      "$LIST" "$WORK/report-$leg.txt"
  )
done

sha256sum "$WORK"/target-*/release/opc-open-oracle > "$HERE/binaries.sha256"
sha256sum "$WORK"/report-*.txt > "$HERE/reports.sha256"
python3 "$HERE/summarize.py" "$WORK/report-before.txt" "$WORK/report-after.txt" "$HERE"
wc -l "$WORK"/report-*.txt
printf 'oracle_work=%s ' "$WORK"
if cmp -s "$WORK/report-before.txt" "$WORK/report-after.txt"; then
  echo 'cmp_status=0'
else
  echo 'cmp_status=1'
  diff -u "$WORK/report-before.txt" "$WORK/report-after.txt" | head -40
  exit 1
fi
