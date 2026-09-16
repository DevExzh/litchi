#!/usr/bin/env bash
# Reproduce change 0632's read-grammar differential in full.
#
#   run-differential.sh <work-dir>
#
# Change 0632 is a read-path change to `soapberry-zip`, so the gate change 0587
# names for any read-grammar change is change 0582's differential.  That harness
# exercises only the strict-layout APIs, so this packet runs change 0611's
# extended copy verbatim: `read_grammar_differential.rs` is change 0582's
# harness plus `I.read_entry` (the ordinary indexed read path) and `R.read`
# (the slice-backed control).  `indexed.open` is the verdict line this change
# can move: `IndexedArchive::from_reader_with_limits` is the constructor whose
# locate and central-directory scan are changed.
#
# Two trees, each with its own CARGO_TARGET_DIR.  Nothing is built against the
# shared working tree and nothing is built inside the change's own worktree.
#
#   before   `git archive` of BEFORE_REV, unmodified.
#   after    the same extraction with change 0632's three production files
#            overlaid from WORKTREE.
#
# The corpus and its generator are change 0582's, unchanged: RNG_SEED
# 0x05820580 regenerates the 22,875-input corpus byte for byte.
set -euo pipefail

HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
REPO=/home/zhuhe/code/litchi
WORKTREE=/home/zhuhe/code/litchi-worktrees/0632
R0582="$REPO/docs/performance/results/change-0582"
WORK="${1:?usage: run-differential.sh <work-dir>}"
BEFORE_REV=c7326f680
JOBS=6

mkdir -p "$WORK"/{before,after}

echo "=== stage 1/7: extracting both trees from $BEFORE_REV ==="
for tree in before after; do
  git -C "$REPO" archive --format=tar "$BEFORE_REV" \
    | tar -x --exclude='docs/*' --exclude='media/*' -C "$WORK/$tree"
  git -C "$REPO" archive --format=tar "$BEFORE_REV" \
      docs/performance/results/change-0416/corpus | tar -x -C "$WORK/$tree"
  cp "$REPO/Cargo.lock" "$WORK/$tree/Cargo.lock"
done

echo "=== stage 2/7: overlaying change 0632's three production files into after/ ==="
for f in locator.rs archive.rs office.rs; do
  cp "$WORKTREE/crates/soapberry-zip/src/$f" "$WORK/after/crates/soapberry-zip/src/$f"
done
for tree in before after; do
  mkdir -p "$WORK/$tree/crates/soapberry-zip/examples"
  cp "$HERE/read_grammar_differential.rs" "$WORK/$tree/crates/soapberry-zip/examples/"
done
echo "before/after production diff:"
for f in locator.rs archive.rs office.rs; do
  diff -q "$WORK/before/crates/soapberry-zip/src/$f" \
          "$WORK/after/crates/soapberry-zip/src/$f" || true
done

echo "=== stage 3/7: regenerating change 0582's corpus ==="
rm -rf "$WORK/corpus"
python3 "$R0582/build_corpus.py" "$REPO" "$WORK/corpus"

echo "=== stage 4/7: building both examples, one cargo process at a time ==="
for tree in before after; do
  echo "--- build $tree ---"
  ( cd "$WORK/$tree" && CARGO_TARGET_DIR="$WORK/target-$tree" \
      nice -n 5 cargo build --release --locked -p soapberry-zip \
        --example read_grammar_differential -j "$JOBS" )
done
sha256sum "$WORK/target-before/release/examples/read_grammar_differential" \
          "$WORK/target-after/release/examples/read_grammar_differential" \
  | tee "$WORK/binaries.sha256"

echo "=== stage 5/7: running the before harness ==="
taskset -c 8 "$WORK/target-before/release/examples/read_grammar_differential" \
    "$WORK/corpus" "$WORK/report-before.txt"

echo "=== stage 6/7: running the after harness ==="
taskset -c 8 "$WORK/target-after/release/examples/read_grammar_differential" \
    "$WORK/corpus" "$WORK/report-after.txt"

echo "=== stage 7/7: classifying ==="
python3 "$HERE/classify.py" "$WORK/report-before.txt" "$WORK/report-after.txt" \
    "$WORK/divergences.tsv" > "$WORK/classification-summary.json"

echo
echo "reports:    $WORK/report-before.txt  $WORK/report-after.txt"
echo "summary:    $WORK/classification-summary.json"
echo "divergences:$WORK/divergences.tsv"
echo "binaries:   $WORK/binaries.sha256"
