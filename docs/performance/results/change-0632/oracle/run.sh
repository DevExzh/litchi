#!/usr/bin/env bash
# Change 0632's open-differential oracle: one report per leg, compared with cmp.
#
# Usage: run.sh <work-dir> <file-list>
set -euo pipefail
HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
WORK="${1:?usage: run.sh <work-dir> <file-list>}"
LIST="${2:?usage: run.sh <work-dir> <file-list>}"
BEFORE_TREE=/home/zhuhe/code/litchi-worktrees/before-c7326f680
AFTER_TREE=/home/zhuhe/code/litchi-worktrees/0632
mkdir -p "$WORK"
for leg in before after; do
  case "$leg" in
    before) TREE="$BEFORE_TREE" ;;
    after)  TREE="$AFTER_TREE" ;;
  esac
  mkdir -p "$WORK/oracle-$leg/src"
  cp "$HERE/src/main.rs" "$WORK/oracle-$leg/src/main.rs"
  sed "s#@TREE@#$TREE#g" "$HERE/Cargo.toml.template" > "$WORK/oracle-$leg/Cargo.toml"
  echo "== building $leg oracle against $TREE"
  ( cd "$WORK/oracle-$leg" && CARGO_TARGET_DIR="$WORK/target-$leg" nice -n 5 \
      cargo build --release -j 6 >/dev/null )
  echo "== running $leg"
  ( cd "$AFTER_TREE" && taskset -c 8 "$WORK/target-$leg/release/opc-open-oracle" \
      "$LIST" "$WORK/report-$leg.txt" "${3:-}" )
done
sha256sum "$WORK"/target-*/release/opc-open-oracle | tee "$WORK/binaries.sha256"
wc -l "$WORK"/report-*.txt
if cmp -s "$WORK/report-before.txt" "$WORK/report-after.txt"; then
  echo "IDENTICAL"
else
  echo "DIVERGENT"
  diff "$WORK/report-before.txt" "$WORK/report-after.txt" | head -40
fi
