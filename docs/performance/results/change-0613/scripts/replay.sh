#!/usr/bin/env bash
# Replay change 0613's measurements from a clean checkout.
#
# Both legs build the same harness from the same workspace; the only
# difference is whether patch/0613-original-audit-memo.patch is applied.
# Every measured process is pinned to one CPU; deterministic counts run
# before timing.
set -euo pipefail

BASE=2d6fbeaed2083de104bbb0b52b2990ce69ac7274
CPU=${CPU:-9}
CASE=xlsx_source_backed_cell_values_one_edit_save
HERE=$(cd "$(dirname "$0")/.." && pwd)
WORK=${WORK:?set WORK to a scratch directory on disk, never under /tmp}

# --- both legs -------------------------------------------------------------
git worktree add --detach "$WORK/before" "$BASE"
git worktree add --detach "$WORK/after"  "$BASE"
git -C "$WORK/after" apply "$HERE/patch/0613-original-audit-memo.patch"

for leg in before after; do
  CARGO_TARGET_DIR="$WORK/target-$leg" \
    cargo build --release --locked \
      --manifest-path "$WORK/$leg/tools/perf-baseline/Cargo.toml" \
      --bin litchi-perf-baseline
done
BEFORE_BIN=$WORK/target-before/release/litchi-perf-baseline
AFTER_BIN=$WORK/target-after/release/litchi-perf-baseline
sha256sum "$BEFORE_BIN" "$AFTER_BIN"

# --- deterministic counts: one callgrind isolation pair per leg -------------
mkdir -p "$WORK/cg"
for leg in before after; do
  case $leg in before) BIN=$BEFORE_BIN ;; after) BIN=$AFTER_BIN ;; esac
  for n in 1 3; do
    taskset -c "$CPU" valgrind --tool=callgrind \
      --callgrind-out-file="$WORK/cg/$leg-n$n.out" --cache-sim=no --branch-sim=no \
      "$BIN" --case "$CASE" --warmup 0 --samples "$n" > "$WORK/cg/$leg-n$n.json"
    python3 "$HERE/scripts/audit-call-counts.py" "$WORK/cg/$leg-n$n.out" \
      xml_minifier::audit::verify_authored \
      xml_minifier::audit::package::is_xml_part
    grep -m1 '^summary:' "$WORK/cg/$leg-n$n.out"
    callgrind_annotate --inclusive=yes --threshold=100 "$WORK/cg/$leg-n$n.out" \
      | grep -E 'write_topology_to_stream \[|verify_authored \[|validate_overlay_xml \[|validate_original_part_xml \['
  done
done
# Per operation = (n3 - n1) / 2: one iteration publishes both corpus shapes.

# --- paired timing, A1 B1 B2 A2 plus the A/A floor --------------------------
python3 "$HERE/scripts/paired-timing.py" \
  "$BEFORE_BIN" "$AFTER_BIN" "$CPU" "$CASE" 30 "$WORK/timing-one-edit.json"

# --- gates on the candidate -------------------------------------------------
cd "$WORK/after"
cargo fmt --all --check
cargo clippy -p litchi-opc -p xml-minifier --all-targets
cargo test -p litchi-opc -p xml-minifier
cargo doc -p litchi-opc -p xml-minifier --no-deps
cargo test -p litchi-xlsx -p litchi-docx -p litchi-pptx
