#!/usr/bin/env bash
# Baseline every selector change 0664 adds, plus the generated-corpus
# selectors they are read against.
#
# Usage: run_baseline.sh <worktree> <staging-dir> <output-dir> <cpu>
#
# The harness is the only thing this change touches: `crates/` is byte
# identical to the base commit, so a run from this worktree measures the base
# library.  Binaries are staged outside every Cargo target directory (change
# 0627's lesson) and every measured process is pinned to one CPU.
set -euo pipefail

WORKTREE="${1:?worktree}"
STAGE="${2:?staging dir}"
OUT="${3:?output dir}"
CPU="${4:?cpu}"

mkdir -p "$STAGE" "$OUT" "$STAGE/fs-root"

PPTX_SAVE="pptx_marker_ordinary_save_lifecycle,pptx_marker_ordinary_save_edit,pptx_marker_ordinary_save_atomic_publish,pptx_marker_ordinary_save_counting_publish,pptx_marker_control_ordinary_save_lifecycle,pptx_marker_control_ordinary_save_edit,pptx_marker_control_ordinary_save_atomic_publish,pptx_marker_control_ordinary_save_counting_publish"
DOCX_SAVE="docx_marker_ordinary_save_lifecycle,docx_marker_ordinary_save_edit,docx_marker_ordinary_save_atomic_publish,docx_marker_ordinary_save_counting_publish,docx_marker_control_ordinary_save_lifecycle,docx_marker_control_ordinary_save_edit,docx_marker_control_ordinary_save_atomic_publish,docx_marker_control_ordinary_save_counting_publish"
READ="pptx_marker_eager_full_text,pptx_marker_control_eager_full_text,pptx_marker_source_full_text,pptx_marker_control_source_full_text,docx_marker_eager_full_text,docx_marker_control_eager_full_text,docx_marker_source_full_text,docx_marker_control_source_full_text"
GENERATED="pptx_ordinary_save_lifecycle,pptx_ordinary_save_edit,pptx_ordinary_save_atomic_publish,pptx_ordinary_save_counting_publish,docx_ordinary_save_lifecycle,docx_ordinary_save_edit,docx_ordinary_save_atomic_publish,docx_ordinary_save_counting_publish"
SINK="docx_semantic_text_to_sink,docx_source_text_to_sink"

build() {
  ( cd "$WORKTREE/tools/perf-baseline" && cargo build --release --locked --bin litchi-perf-baseline )
  ( cd "$WORKTREE/tools/perf-baseline" && cargo build --release --locked --features allocator-metrics --bin litchi-perf-baseline-alloc )
  cp "$WORKTREE/tools/perf-baseline/target/release/litchi-perf-baseline" "$STAGE/litchi-perf-baseline"
  cp "$WORKTREE/tools/perf-baseline/target/release/litchi-perf-baseline-alloc" "$STAGE/litchi-perf-baseline-alloc"
  sha256sum "$STAGE/litchi-perf-baseline" "$STAGE/litchi-perf-baseline-alloc" > "$OUT/binary-sha256.txt"
}

timed() { # <group> <cases> <repeat> [extra...]
  local group="$1" cases="$2" repeat="$3"; shift 3
  taskset -c "$CPU" "$STAGE/litchi-perf-baseline" \
    --warmup 20 --samples 50 \
    --filesystem-root "$STAGE/fs-root" \
    --case "$cases" "$@" \
    --json "$OUT/${group}-${repeat}.json" > "$OUT/${group}-${repeat}.txt" 2>&1
}

allocated() { # <group> <cases> [extra...]
  local group="$1" cases="$2"; shift 2
  taskset -c "$CPU" "$STAGE/litchi-perf-baseline-alloc" \
    --warmup 5 --samples 20 \
    --filesystem-root "$STAGE/fs-root" \
    --case "$cases" "$@" \
    --json "$OUT/alloc-${group}.json" > "$OUT/alloc-${group}.txt" 2>&1
}

build

# Deterministic evidence first: the per-member census of every marker corpus.
taskset -c "$CPU" "$STAGE/litchi-perf-baseline" \
  --warmup 0 --samples 1 --filesystem-root "$STAGE/fs-root" \
  --case "$READ" --marker-evidence "$OUT/marker-census.json" \
  --json "$OUT/census-run.json" > "$OUT/census-run.txt" 2>&1

# Allocation counts (deterministic; the allocator observer perturbs timing, so
# these runs are never used for latency).
allocated pptx-save "$PPTX_SAVE"
allocated docx-save "$DOCX_SAVE"
allocated generated-save "$GENERATED"
allocated read "$READ"
allocated sink "$SINK" --semantic-shape medium

# Timing last, three repeats, ordered A1 B1 B2 A2 within each group by the
# harness's own case ordering.
for repeat in R1 R2 R3; do
  timed pptx-save "$PPTX_SAVE" "$repeat"
  timed docx-save "$DOCX_SAVE" "$repeat"
  timed read "$READ" "$repeat"
  timed generated-save "$GENERATED" "$repeat"
  timed sink "$SINK" "$repeat" --semantic-shape medium
done

# A/A floor: two further runs of the same group in the same window.
timed pptx-save-aa "$PPTX_SAVE" A1
timed pptx-save-aa "$PPTX_SAVE" A2
timed read-aa "$READ" A1
timed read-aa "$READ" A2

echo "done"
