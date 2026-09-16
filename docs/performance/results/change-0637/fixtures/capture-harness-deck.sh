#!/usr/bin/env bash
# Capture the 200-slide PPTX deck `tools/perf-baseline` builds for its
# filesystem query controls (`build_pptx_source_edit_corpus`: 200 slides,
# 8 text boxes each, 8 x 2 MiB deterministic PNGs).
#
# The harness materializes the corpus under `--filesystem-root` and removes the
# run directory when it finishes, and it exposes no corpus-export command, so
# this copies the file out while a long run holds it. The deck is 17,017,139
# bytes, sha256
# 61b2b99083ca27ebd37955db600955e3f41289b93dba71951983164239eff757; it is not
# retained in this packet because it is reproducible and large.
set -euo pipefail
OUT="${1:?usage: capture-harness-deck.sh <output.pptx>}"
BIN="${2:?usage: capture-harness-deck.sh <output.pptx> <litchi-perf-baseline>}"
ROOT=$(mktemp -d)
( cd "$(dirname "$BIN")" && taskset -c 13 "$BIN" --case pptx_file_eager_slide_count \
    --samples 200 --warmup 3 --filesystem-root "$ROOT" >/dev/null 2>&1 ) &
RUNPID=$!
for _ in $(seq 1 1200); do
  f=$(find "$ROOT" -name '*.pptx' -size +1k 2>/dev/null | head -1)
  if [ -n "$f" ]; then cp "$f" "$OUT"; break; fi
  sleep 0.25
done
kill $RUNPID 2>/dev/null || true
wait $RUNPID 2>/dev/null || true
rm -rf "$ROOT"
sha256sum "$OUT"
